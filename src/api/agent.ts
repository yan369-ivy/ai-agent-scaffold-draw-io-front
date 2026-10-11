import { API_CONFIG } from '@/config/api-config';
import { 
    Response, 
    AiAgentConfigResponseDTO, 
    CreateSessionResponseDTO, 
    ChatRequestDTO,
    ChatResponseDTO 
} from '@/types/api';

const handleResponse = async <T>(response: globalThis.Response): Promise<Response<T>> => {
    if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`HTTP error! status: ${response.status}, message: ${errorText}`);
    }
    const data = await response.json();
    if (data.code !== "0000") {
        throw new Error(data.info || 'Unknown API error');
    }
    return data;
};

// Types for streaming drawio events
export interface DrawioNodeChunk {
    type: 'drawio_node';
    id: string;
    label: string;
    xml: string;
}

export interface DrawioEdgeChunk {
    type: 'drawio_edge';
    id: string;
    label: string;
    source: string;
    target: string;
    xml: string;
}

export interface DrawioDoneChunk {
    type: 'drawio_done';
    content: string;
}

export interface DrawioLegacyChunk {
    type: 'drawio';
    content: string;
}

export interface StatusChunk {
    type: 'status';
    content: string;
}

export interface ErrorChunk {
    type: 'error';
    content: string;
}

export interface UserChunk {
    type: 'user';
    content: string;
}

export interface RawTextChunk {
    type: 'raw_text';
    content: string;
}

export interface DoneChunk {
    type: 'done';
}

// PPT streaming chunks
export interface PptSlideChunk {
    type: 'ppt_slide';
    slide: any;
    title?: string;
    slideIndex?: number;
}

export interface PptDoneChunk {
    type: 'ppt_done';
    title?: string;
    totalSlides?: number;
}

export interface PptRawChunk {
    type: 'ppt_raw';
    raw: string;
}

export interface TokenChunk {
    type: 'token';
    content: string;
}

export type StreamChunk = DrawioNodeChunk | DrawioEdgeChunk | DrawioDoneChunk | DrawioLegacyChunk | StatusChunk | ErrorChunk | UserChunk | RawTextChunk | DoneChunk | PptSlideChunk | PptDoneChunk | PptRawChunk | TokenChunk;

export interface StreamEvent {
    phase: 'analyzing' | 'drawing' | 'reviewing' | 'thinking' | 'error' | 'done' | 'generating';
    chunk: StreamChunk;
}

export type StreamEventCallback = (event: StreamEvent) => void;

export const agentApi = {
    /**
     * Query AI Agent Config List
     * Path: /api/v1/query_ai_agent_config_list
     */
    queryAiAgentConfigList: async (): Promise<Response<AiAgentConfigResponseDTO[]>> => {
        const response = await fetch(`${API_CONFIG.BASE_URL}/query_ai_agent_config_list`, {
            method: 'GET',
            headers: {
                'Content-Type': 'application/json',
            },
        });
        return handleResponse<AiAgentConfigResponseDTO[]>(response);
    },

    /**
     * Create Session
     * Path: /api/v1/create_session
     */
    createSession: async (agentId: string, userId: string): Promise<Response<CreateSessionResponseDTO>> => {
        const response = await fetch(`${API_CONFIG.BASE_URL}/create_session`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({ agentId, userId }),
        });
        return handleResponse<CreateSessionResponseDTO>(response);
    },

    /**
     * Chat (blocking)
     * Path: /api/v1/chat
     */
    chat: async (data: ChatRequestDTO): Promise<Response<ChatResponseDTO>> => {
        const response = await fetch(`${API_CONFIG.BASE_URL}/chat`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(data),
        });
        return handleResponse<ChatResponseDTO>(response);
    },

    /**
     * Chat Stream (SSE-like streaming)
     * Path: /api/v1/chat_stream
     * Accepts newline-delimited structured events or plain text from the scaffold backend.
     */
    chatStream: async (
        data: ChatRequestDTO,
        onEvent: StreamEventCallback,
        onError: (error: Error) => void,
        onComplete: () => void
    ): Promise<AbortController> => {
        const controller = new AbortController();

        try {
            const response = await fetch(`${API_CONFIG.BASE_URL}/chat_stream`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify(data),
                signal: controller.signal,
            });

            if (!response.ok) {
                const errorText = await response.text();
                onError(new Error(`HTTP error! status: ${response.status}, message: ${errorText}`));
                return controller;
            }

            const reader = response.body!.getReader();
            const decoder = new TextDecoder();
            let buffer = '';
            let mode: 'detect' | 'structured' | 'plain' = 'detect';

            const parseEvent = (line: string): StreamEvent | null => {
                try {
                    const value: unknown = JSON.parse(line);
                    if (value && typeof value === 'object' && 'phase' in value && 'chunk' in value &&
                        typeof value.phase === 'string' && value.chunk && typeof value.chunk === 'object' &&
                        'type' in value.chunk && typeof value.chunk.type === 'string') {
                        return value as StreamEvent;
                    }
                } catch {
                    // Plain text may also start with JSON-like content.
                }
                return null;
            };

            const emitPlainText = (content: string) => {
                if (content) onEvent({ phase: 'thinking', chunk: { type: 'raw_text', content } });
            };

            const processStructuredLines = () => {
                let lineEnd = buffer.indexOf('\n');
                while (lineEnd !== -1) {
                    const line = buffer.slice(0, lineEnd).trim();
                    buffer = buffer.slice(lineEnd + 1);
                    if (line) {
                        const event = parseEvent(line);
                        if (event) onEvent(event);
                        else console.warn('Failed to parse stream event:', line);
                    }
                    lineEnd = buffer.indexOf('\n');
                }
            };

            const processText = (text: string) => {
                if (mode === 'plain') {
                    emitPlainText(text);
                    return;
                }

                buffer += text;
                if (mode === 'detect') {
                    const candidate = buffer.trimStart();
                    if (!candidate) return;
                    if (!candidate.startsWith('{')) {
                        mode = 'plain';
                        emitPlainText(buffer);
                        buffer = '';
                        return;
                    }

                    const lineEnd = candidate.indexOf('\n');
                    if (lineEnd === -1) return;
                    if (!parseEvent(candidate.slice(0, lineEnd).trim())) {
                        mode = 'plain';
                        emitPlainText(buffer);
                        buffer = '';
                        return;
                    }
                    mode = 'structured';
                    buffer = candidate;
                }
                processStructuredLines();
            };

            const processStream = async () => {
                try {
                    while (true) {
                        const { done, value } = await reader.read();
                        if (done) break;
                        processText(decoder.decode(value, { stream: true }));
                    }

                    processText(decoder.decode());
                    if (mode === 'detect') {
                        const event = parseEvent(buffer.trim());
                        if (event) onEvent(event);
                        else emitPlainText(buffer);
                    } else if (mode === 'structured' && buffer.trim()) {
                        const event = parseEvent(buffer.trim());
                        if (event) onEvent(event);
                        else console.warn('Failed to parse stream event:', buffer);
                    }

                    // Only call onComplete if we didn't abort
                    if (!controller.signal.aborted) {
                        onComplete();
                    }
                } catch (err: any) {
                    if (err.name === 'AbortError') {
                        // User cancelled, no error, but call complete to cleanup UI state
                        onComplete();
                        return;
                    }
                    onError(err instanceof Error ? err : new Error(String(err)));
                }
            };

            processStream();
        } catch (err: any) {
            if (err.name !== 'AbortError') {
                onError(err instanceof Error ? err : new Error(String(err)));
            }
        }

        return controller;
    }
};
