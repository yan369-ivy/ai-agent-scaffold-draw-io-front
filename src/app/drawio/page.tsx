'use client';

import { DrawIoEmbed, DrawIoEmbedRef } from 'react-drawio';
import { useRef, useState, useEffect } from 'react';
import { useRouter } from 'next/navigation';
import { getUserInfo, clearUserInfo } from '@/utils/cookie';
import { agentApi, StreamEvent, DrawioNodeChunk, DrawioEdgeChunk, DrawioDoneChunk, DrawioLegacyChunk, UserChunk, StatusChunk, ErrorChunk } from '@/api/agent';
import { AiAgentConfigResponseDTO } from '@/types/api';
import ReactMarkdown from 'react-markdown';
import remarkGfm from 'remark-gfm';

// Message type definition
type MessageStep = {
  phase: string;
  label: string;
  content: string;
  status: 'running' | 'done' | 'pending';
};

type Message = {
  id: string;
  role: 'user' | 'agent';
  content: string;
  reasoning?: string;
  steps?: MessageStep[];
  timestamp: number;
};

// Elegant SVG Icons with consistent styling
const Icons = {
  Chat: ({ className }: { className?: string }) => (
    <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}>
      <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"></path>
    </svg>
  ),
  Close: ({ className }: { className?: string }) => (
    <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}>
      <line x1="18" y1="6" x2="6" y2="18"></line>
      <line x1="6" y1="6" x2="18" y2="18"></line>
    </svg>
  ),
  Send: ({ className }: { className?: string }) => (
    <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}>
      <line x1="22" y1="2" x2="11" y2="13"></line>
      <polygon points="22 2 15 22 11 13 2 9 22 2"></polygon>
    </svg>
  ),
  User: ({ className }: { className?: string }) => (
    <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}>
      <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"></path>
      <circle cx="12" cy="7" r="4"></circle>
    </svg>
  ),
  Bot: ({ className }: { className?: string }) => (
    <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}>
      <path d="M12 2a2 2 0 0 1 2 2v2a2 2 0 0 1-2 2 2 2 0 0 1-2-2V4a2 2 0 0 1 2-2z"></path>
      <path d="M4 11v6a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-6a2 2 0 0 0-2-2H6a2 2 0 0 0-2 2z"></path>
      <path d="M9 22v-3"></path>
      <path d="M15 22v-3"></path>
    </svg>
  ),
  Square: ({ className }: { className?: string }) => (
    <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="currentColor" className={className}>
      <rect x="6" y="6" width="12" height="12" rx="2" ry="2"></rect>
    </svg>
  ),
  Download: ({ className }: { className?: string }) => (
    <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}>
      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
      <polyline points="7 10 12 15 17 10"></polyline>
      <line x1="12" y1="15" x2="12" y2="3"></line>
    </svg>
  ),
  Sparkles: ({ className }: { className?: string }) => (
    <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}>
      <path d="m12 3-1.912 5.813a2 2 0 0 1-1.275 1.275L3 12l5.813 1.912a2 2 0 0 1 1.275 1.275L12 21l1.912-5.813a2 2 0 0 1 1.275-1.275L21 12l-5.813-1.912a2 2 0 0 1-1.275-1.275L12 3Z" />
    </svg>
  ),
  Logout: ({ className }: { className?: string }) => (
    <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}>
      <path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"></path>
      <polyline points="16 17 21 12 16 7"></polyline>
      <line x1="21" y1="12" x2="9" y2="12"></line>
    </svg>
  ),
  Layers: ({ className }: { className?: string }) => (
    <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}>
      <polygon points="12 2 2 7 12 12 22 7 12 2"></polygon>
      <polyline points="2 17 12 22 22 17"></polyline>
      <polyline points="2 12 12 17 22 12"></polyline>
    </svg>
  ),
  Loader: ({ className }: { className?: string }) => (
    <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={`animate-spin ${className}`}>
      <path d="M21 12a9 9 0 1 1-6.219-8.56"></path>
    </svg>
  ),
  Plus: ({ className }: { className?: string }) => (
    <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}>
      <line x1="12" y1="5" x2="12" y2="19"></line>
      <line x1="5" y1="12" x2="19" y2="12"></line>
    </svg>
  ),
  Trash: ({ className }: { className?: string }) => (
    <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}>
      <polyline points="3 6 5 6 21 6"></polyline>
      <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
    </svg>
  ),
  MessageSquare: ({ className }: { className?: string }) => (
    <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}>
      <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"></path>
    </svg>
  )
};

interface Session {
  id: string;
  backendSessionId?: string;
  title: string;
  messages: Message[];
  drawIoXml: string | null;
  lastModified: number;
}

export interface CustomModelConfig {
  id: string;
  name: string;
  baseUrl: string;
  apiKey: string;
  model: string;
  completionsPath: string;
  enabled: boolean;
}

export default function Home() {
  const router = useRouter();
  const [imgData, setImgData] = useState<string | null>(null);
  const drawioRef = useRef<DrawIoEmbedRef>(null);
  
  // User State
  const [currentUser, setCurrentUser] = useState('');

  // Chat State
  const [isChatOpen, setIsChatOpen] = useState(true);
  const [messages, setMessages] = useState<Message[]>([
    {
      id: '1',
      role: 'agent',
      content: '你好！我是你的智能架构助手。请选择一个智能体开始对话。',
      timestamp: Date.now()
    }
  ]);
  const [inputValue, setInputValue] = useState('');
  const [isSending, setIsSending] = useState(false);
  const messagesEndRef = useRef<HTMLDivElement>(null);

  // Stream State
  const [streamPhase, setStreamPhase] = useState<string>('');
  const [streamProgress, setStreamProgress] = useState<string>('');
  const streamAbortRef = useRef<AbortController | null>(null);

  // Context State
  const [useHistoryContext, setUseHistoryContext] = useState(false);
  const [lastExportedData, setLastExportedData] = useState<{data: string, timestamp: number} | null>(null);
  const isExportingForChatRef = useRef(false);
  const isAutosaveRef = useRef(false);
  const pendingMessageRef = useRef('');
  const [isDrawIoReady, setIsDrawIoReady] = useState(false);
  const initialLoadDoneRef = useRef(false);

  // Agent State
  const [agents, setAgents] = useState<AiAgentConfigResponseDTO[]>([]);
  const [selectedAgentId, setSelectedAgentId] = useState('');
  const [sessionId, setSessionId] = useState('');

  // Rename State
  const [isRenameModalOpen, setIsRenameModalOpen] = useState(false);
  const [renamingSessionId, setRenamingSessionId] = useState<string | null>(null);
  const [newSessionTitle, setNewSessionTitle] = useState('');

  // Custom API Config State
  const [showApiConfig, setShowApiConfig] = useState(false);
  const [customModels, setCustomModels] = useState<CustomModelConfig[]>([]);
  const [selectedCustomModelId, setSelectedCustomModelId] = useState<string>('default');
  
  // Temporary state for editing in modal
  const [editingModel, setEditingModel] = useState<CustomModelConfig | null>(null);

  // Session Management State
  const [sessions, setSessions] = useState<Session[]>([]);
  const [currentSessionId, setCurrentSessionId] = useState<string | null>(null);
  const currentSessionRef = useRef(currentSessionId);

  // Update ref
  useEffect(() => {
    currentSessionRef.current = currentSessionId;
  }, [currentSessionId]);

  // Handle Initial Load
  useEffect(() => {
    if (!initialLoadDoneRef.current && isDrawIoReady && currentSessionId && sessions.length > 0) {
      const session = sessions.find(s => s.id === currentSessionId);
      if (session && session.drawIoXml && drawioRef.current) {
        drawioRef.current.load({ xml: session.drawIoXml });
      }
      initialLoadDoneRef.current = true;
    }
  }, [isDrawIoReady, currentSessionId, sessions]);

  // Load sessions from localStorage
  useEffect(() => {
    const savedSessions = localStorage.getItem('drawio_sessions');
    if (savedSessions) {
      try {
        const parsed = JSON.parse(savedSessions);
        setSessions(parsed);
        if (parsed.length > 0) {
          // Load the most recent session (first one if sorted by lastModified desc)
          const mostRecent = parsed.sort((a: Session, b: Session) => b.lastModified - a.lastModified)[0];
          setCurrentSessionId(mostRecent.id);
          setMessages(mostRecent.messages);
          // Note: Draw.io XML loading happens after drawioRef is ready or when we switch
        } else {
            createNewSession(true);
        }
      } catch (e) {
        console.error('Failed to parse sessions:', e);
        createNewSession(true);
      }
    } else {
      createNewSession(true);
    }
  }, []);

  // Save sessions to localStorage whenever they change
  useEffect(() => {
    if (sessions.length > 0) {
      try {
        localStorage.setItem('drawio_sessions', JSON.stringify(sessions));
      } catch (e) {
        console.error('Failed to save sessions to localStorage:', e);
      }
    }
  }, [sessions]);

  // Update current session messages and backendSessionId when they change
  useEffect(() => {
    if (currentSessionId) {
      setSessions(prev => prev.map(session => {
        if (session.id === currentSessionId) {
          return {
            ...session,
            messages,
            backendSessionId: sessionId,
            // Update title if it's the default "New Chat" and we have a user message
            title: session.title === 'New Chat' && messages.find(m => m.role === 'user') 
              ? (messages.find(m => m.role === 'user')?.content.slice(0, 20) || 'New Chat')
              : session.title
          };
        }
        return session;
      }));
    }
  }, [messages, currentSessionId, sessionId]);

  const createNewSession = (isInitial = false, backendId = '') => {
    const newSession: Session = {
      id: Date.now().toString(),
      backendSessionId: backendId,
      title: 'New Chat',
      messages: [{
        id: Date.now().toString(),
        role: 'agent',
        content: '你好！我是你的智能架构助手。请选择一个智能体开始对话。',
        timestamp: Date.now()
      }],
      drawIoXml: null,
      lastModified: Date.now()
    };

    setSessions(prev => [newSession, ...prev]);
    setCurrentSessionId(newSession.id);
    setMessages(newSession.messages);
    setSessionId(backendId);
    
    if (!isInitial && drawioRef.current) {
      drawioRef.current.load({ xml: '' }); // Clear diagram
    }
  };

  const handleSwitchSession = (targetSessionId: string) => {
    if (targetSessionId === currentSessionId) return;
    loadSession(targetSessionId);
  };

  const loadSession = (targetSessionId: string) => {
    const session = sessions.find(s => s.id === targetSessionId);
    if (session) {
        setCurrentSessionId(targetSessionId);
        setMessages(session.messages);
        setSessionId(session.backendSessionId || '');
        if (drawioRef.current && session.drawIoXml) {
            drawioRef.current.load({ xml: session.drawIoXml });
        } else if (drawioRef.current) {
            drawioRef.current.load({ xml: '' });
        }
    }
  };

  const handleDeleteSession = (e: React.MouseEvent, sessionIdToDelete: string) => {
    e.stopPropagation();
    const newSessions = sessions.filter(s => s.id !== sessionIdToDelete);
    setSessions(newSessions);
    localStorage.setItem('drawio_sessions', JSON.stringify(newSessions));

    if (currentSessionId === sessionIdToDelete) {
        if (newSessions.length > 0) {
            loadSession(newSessions[0].id);
        } else {
            createNewSession();
        }
    }
  };

  const handleDoubleClickSession = (session: Session) => {
    setRenamingSessionId(session.id);
    setNewSessionTitle(session.title);
    setIsRenameModalOpen(true);
  };

  const handleRenameSave = () => {
    if (renamingSessionId && newSessionTitle.trim()) {
      setSessions(prev => prev.map(s => 
        s.id === renamingSessionId ? { ...s, title: newSessionTitle.trim() } : s
      ));
      setIsRenameModalOpen(false);
      setRenamingSessionId(null);
      setNewSessionTitle('');
    }
  };

  const saveCustomModels = (models: CustomModelConfig[]) => {
    setCustomModels(models);
    localStorage.setItem('ai_agent_custom_models', JSON.stringify(models));
  };

  const handleAddNewModel = () => {
    setEditingModel({
      id: Date.now().toString(),
      name: '新模型',
      baseUrl: 'https://api.openai.com',
      apiKey: '',
      model: 'gpt-4o',
      completionsPath: 'v1/chat/completions',
      enabled: true
    });
  };

  const handleSaveEditingModel = () => {
    if (!editingModel) return;
    const exists = customModels.some(m => m.id === editingModel.id);
    let newModels;
    if (exists) {
      newModels = customModels.map(m => m.id === editingModel.id ? editingModel : m);
    } else {
      newModels = [...customModels, editingModel];
    }
    saveCustomModels(newModels);
    setSelectedCustomModelId(editingModel.id);
    localStorage.setItem('ai_agent_selected_model', editingModel.id);
    setEditingModel(null);
  };

  const handleDeleteModel = (id: string) => {
    const newModels = customModels.filter(m => m.id !== id);
    saveCustomModels(newModels);
    if (selectedCustomModelId === id) {
      setSelectedCustomModelId('default');
      localStorage.setItem('ai_agent_selected_model', 'default');
    }
  };

  const exportDiagram = () => {
    if (drawioRef.current) {
      drawioRef.current.exportDiagram({
        format: 'xmlsvg'
      });
    }
  };

  const scrollToBottom = () => {
    messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
  };

  useEffect(() => {
    scrollToBottom();
  }, [messages, isChatOpen]);

  // Check Login & Load Agents
  useEffect(() => {
    const userInfo = getUserInfo();
    if (!userInfo || !userInfo.user) {
      router.push('/login');
      return;
    }
    setCurrentUser(userInfo.user);

    // Load Custom Models
    const savedModels = localStorage.getItem('ai_agent_custom_models');
    if (savedModels) {
      try {
        setCustomModels(JSON.parse(savedModels));
      } catch (e) {}
    }
    const savedSelected = localStorage.getItem('ai_agent_selected_model');
    if (savedSelected) {
      setSelectedCustomModelId(savedSelected);
    }

    // Load Agents
    const loadAgents = async () => {
      try {
        const res = await agentApi.queryAiAgentConfigList();
        setAgents(res.data || []);
        if (res.data && res.data.length > 0) {
          // Try to restore last agent or default to first
          const lastAgentId = localStorage.getItem('ai_agent_last_agent');
          if (lastAgentId && res.data.find(a => a.agentId === lastAgentId)) {
            setSelectedAgentId(lastAgentId);
          } else {
            setSelectedAgentId(res.data[0].agentId);
          }
        }
      } catch (error) {
        console.error('Failed to load agents:', error);
        setMessages(prev => [...prev, {
          id: Date.now().toString(),
          role: 'agent',
          content: '加载智能体列表失败，请检查后端服务是否启动。',
          timestamp: Date.now()
        }]);
      }
    };
    loadAgents();
  }, [router]);

  const handleLogout = () => {
    clearUserInfo();
    router.push('/login');
  };

  const handleAgentChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const newAgentId = e.target.value;
    setSelectedAgentId(newAgentId);
    setSessionId(''); // Reset session when agent changes
    localStorage.setItem('ai_agent_last_agent', newAgentId);
  };

  const finalizeNewChat = async () => {
    if (!selectedAgentId || !currentUser) return;
    
    try {
        const res = await agentApi.createSession(selectedAgentId, currentUser);
        createNewSession(false, res.data.sessionId);
        setInputValue('');
    } catch (error) {
        console.error('Failed to create new session:', error);
    }
  };

  const handleNewChat = async () => {
     finalizeNewChat();
  };

  const handleRestartSession = async () => {
    if (!selectedAgentId || !currentUser) return;

    if (!currentSessionId) {
      finalizeNewChat();
      return;
    }
    
    try {
        const res = await agentApi.createSession(selectedAgentId, currentUser);
        const newBackendId = res.data.sessionId;
        
        const initialMsg: Message = {
          id: Date.now().toString(),
          role: 'agent',
          content: '你好！我是你的智能架构助手。请选择一个智能体开始对话。',
          timestamp: Date.now()
        };

        setSessionId(newBackendId);
        setMessages([initialMsg]);
        setInputValue('');

        setSessions(prev => prev.map(session => {
          if (session.id === currentSessionId) {
            return {
              ...session,
              backendSessionId: newBackendId,
              messages: [initialMsg],
              lastModified: Date.now()
            };
          }
          return session;
        }));
    } catch (error) {
        console.error('Failed to restart session:', error);
    }
  };

  const handleStopStream = () => {
    if (streamAbortRef.current) {
      streamAbortRef.current.abort();
      streamAbortRef.current = null;
    }
    setIsSending(false);
    setStreamPhase('');
    setStreamProgress('');
    setMessages(prev => [...prev, {
      id: Date.now().toString(),
      role: 'agent',
      content: '⚠️ 已停止生成。',
      timestamp: Date.now()
    }]);
  };

  const performSendMessage = async (displayContent: string, apiContent: string) => {
    if (!selectedAgentId) {
      setMessages(prev => [...prev, {
        id: Date.now().toString(),
        role: 'agent',
        content: '请先选择一个智能体。',
        timestamp: Date.now()
      }]);
      setIsSending(false);
      return;
    }

    setIsSending(true);

    const userMsg: Message = {
      id: Date.now().toString(),
      role: 'user',
      content: displayContent,
      timestamp: Date.now()
    };
    
    const agentMsgId = Date.now().toString() + '-agent';
    const initialAgentMsg: Message = {
      id: agentMsgId,
      role: 'agent',
      content: '',
      reasoning: '',
      steps: [],
      timestamp: Date.now()
    };

    setMessages(prev => [...prev, userMsg, initialAgentMsg]);

    try {
      // 1. Ensure Session
      let activeBackendSessionId = sessionId;
      if (!activeBackendSessionId) {
        const sessionRes = await agentApi.createSession(selectedAgentId, currentUser);
        activeBackendSessionId = sessionRes.data.sessionId;
        setSessionId(activeBackendSessionId);
      }

      // Update session lastModified
      setSessions(prev => prev.map(session => {
        if (session.id === currentSessionId) {
          return { ...session, lastModified: Date.now() };
        }
        return session;
      }));

      // 2. Send Message via Stream
      setStreamPhase('connecting');
      setStreamProgress('正在连接...');

      // Track incremental merge state
      let nodeCount = 0;
      let edgeCount = 0;
      let hasIncrementalContent = false;
      let finalXml = '';
      let agentTextContent = ''; // For non-drawio user-type responses
      let receivedDrawioDone = false;
      let accumulatedNodes: string[] = []; // To hold incrementally added nodes
      let accumulatedEdges: string[] = []; // To hold incrementally added edges
      
      let accumulatedReasoning = '';
      let accumulatedContent = '';
      let accumulatedSteps: MessageStep[] = [];

      // Helper to update steps
      const updateStep = (phaseStr: string, phaseText: string, contentToAdd: string, isDone: boolean = false) => {
          const stepIndex = accumulatedSteps.findIndex(s => s.phase === phaseStr);
          if (stepIndex >= 0) {
              if (contentToAdd) {
                  accumulatedSteps[stepIndex].content += contentToAdd + '\n';
              }
              if (isDone) {
                  accumulatedSteps[stepIndex].status = 'done';
              }
          } else {
              // Mark previous running steps as done
              accumulatedSteps.forEach(s => { if (s.status === 'running') s.status = 'done'; });
              accumulatedSteps.push({
                  phase: phaseStr,
                  label: phaseText,
                  content: contentToAdd ? contentToAdd + '\n' : '',
                  status: isDone ? 'done' : 'running'
              });
          }
      };

      // Clear the canvas before starting a new stream
      if (drawioRef.current && currentSessionId === currentSessionRef.current) {
        try {
          drawioRef.current.load({ xml: '<mxGraphModel><root><mxCell id="0"/><mxCell id="1" parent="0"/></root></mxGraphModel>' });
        } catch (e) {
          console.error('Failed to clear canvas:', e);
        }
      }

      const activeModelConfig = customModels.find(m => m.id === selectedCustomModelId && m.enabled);

      const controller = await agentApi.chatStream(
        {
          agentId: selectedAgentId,
          userId: currentUser,
          sessionId: activeBackendSessionId,
          message: apiContent,
          customBaseUrl: activeModelConfig?.baseUrl || undefined,
          customApiKey: activeModelConfig?.apiKey || undefined,
          customCompletionsPath: activeModelConfig?.completionsPath || undefined,
          customModel: activeModelConfig?.model || undefined
        },
        // onEvent
        (event: StreamEvent) => {
          const { phase, chunk } = event;

          // Update phase display
          const phaseLabel: Record<string, string> = {
            analyzing: '🔍 分析需求',
            drawing: '🎨 绘制图表',
            reviewing: '✅ 检查优化',
            thinking: '🤔 思考中',
          };
          const currentPhaseLabel = phaseLabel[phase] || phaseLabel.thinking;

          if (phase !== 'done' && phase !== 'error') {
            setStreamPhase(phase);
            updateStep(phase, currentPhaseLabel, '');
          }

          switch (chunk.type) {
            case 'drawio_node': {
              // If a previous agent (e.g. drawer) finished, and a new agent (e.g. reviewer) starts sending nodes,
              // we must reset the accumulators so we don't duplicate nodes.
              if (receivedDrawioDone) {
                // Because reviewer sends a FULL list of corrected nodes and edges,
                // we MUST clear the previous accumulated elements so they don't duplicate.
                accumulatedNodes = [];
                accumulatedEdges = [];
                nodeCount = 0;
                edgeCount = 0;
                receivedDrawioDone = false;
                // Preemptively clear the canvas so we can cleanly redraw the reviewer's output
                if (drawioRef.current && currentSessionId === currentSessionRef.current) {
                  try {
                    drawioRef.current.load({ xml: '<mxGraphModel><root><mxCell id="0"/><mxCell id="1" parent="0"/></root></mxGraphModel>' });
                  } catch (e) {
                    console.error('Failed to clear canvas for reviewer:', e);
                  }
                }
              }

              hasIncrementalContent = true;
              nodeCount++;
              setStreamProgress(`添加节点 #${nodeCount}: ${chunk.label}`);

              // Sometimes AI returns empty XML or malformed tags, skip adding to prevent crashing draw.io
              if (chunk.xml && chunk.xml.trim() !== '' && chunk.xml.includes('<mxCell')) {
                  accumulatedNodes.push(chunk.xml);
              }

              // Incrementally load the accumulated cells into draw.io
              if (drawioRef.current && currentSessionId === currentSessionRef.current) {
                try {
                  // Ensure nodes are placed before edges in the XML structure
                  const loadXml = `<mxGraphModel><root><mxCell id="0"/><mxCell id="1" parent="0"/>${accumulatedNodes.join('')}${accumulatedEdges.join('')}</root></mxGraphModel>`;
                  // By avoiding full load and just pushing the xml state, react-drawio updates the iframe
                  drawioRef.current.load({ xml: loadXml });
                } catch (e) {
                  console.error('Failed to load node:', e);
                }
              }
              break;
            }

            case 'drawio_edge': {
              if (receivedDrawioDone) {
                accumulatedNodes = [];
                accumulatedEdges = [];
                nodeCount = 0;
                edgeCount = 0;
                receivedDrawioDone = false;
                if (drawioRef.current && currentSessionId === currentSessionRef.current) {
                  try {
                    drawioRef.current.load({ xml: '<mxGraphModel><root><mxCell id="0"/><mxCell id="1" parent="0"/></root></mxGraphModel>' });
                  } catch (e) {
                    console.error('Failed to clear canvas for reviewer:', e);
                  }
                }
              }

              hasIncrementalContent = true;
              edgeCount++;
              setStreamProgress(`添加连线 #${edgeCount}: ${chunk.label || chunk.source + '→' + chunk.target}`);

              if (chunk.xml && chunk.xml.trim() !== '' && chunk.xml.includes('<mxCell')) {
                  accumulatedEdges.push(chunk.xml);
              }

              // Incrementally load the accumulated cells into draw.io
              if (drawioRef.current && currentSessionId === currentSessionRef.current) {
                try {
                  // Ensure nodes are placed before edges in the XML structure
                  const loadXml = `<mxGraphModel><root><mxCell id="0"/><mxCell id="1" parent="0"/>${accumulatedNodes.join('')}${accumulatedEdges.join('')}</root></mxGraphModel>`;
                  // By avoiding full load and just pushing the xml state, react-drawio updates the iframe
                  drawioRef.current.load({ xml: loadXml });
                } catch (e) {
                  console.error('Failed to load edge:', e);
                }
              }
              break;
            }

            case 'drawio_done': {
              hasIncrementalContent = true;
              receivedDrawioDone = true;
              finalXml = chunk.content;
              setStreamProgress('🎨 绘制完成，正在加载最终图表...');

              // Only render the diagram and end the process if this is the final Reviewer output (or if there is no reviewer)
              // We identify the final output if phase is 'drawing' or 'done' (since we removed reviewer)
              const isFinalStage = phase === 'drawing' || phase === 'done';
              
              if (isFinalStage) {
                  // Final full load to ensure consistency
                  if (drawioRef.current && currentSessionId === currentSessionRef.current && finalXml && finalXml.trim() !== '') {
                    try {
                      drawioRef.current.load({ xml: finalXml });
                    } catch (e) {
                      console.error('Failed to load final diagram:', e);
                    }
                  }

                  // Save final XML to session
                  setSessions(prev => prev.map(session => {
                    if (session.id === currentSessionId) {
                      return {
                        ...session,
                        drawIoXml: finalXml,
                        lastModified: Date.now()
                      };
                    }
                    return session;
                  }));
              }
              
              break;
            }

            case 'drawio': {
              // Legacy format: {"type":"drawio","content":"<xml>"}
              hasIncrementalContent = true;
              receivedDrawioDone = true;
              finalXml = chunk.content;

              const isFinalStage = phase === 'drawing' || phase === 'done';

              if (isFinalStage) {
                  if (drawioRef.current && currentSessionId === currentSessionRef.current && finalXml && finalXml.trim() !== '') {
                    try {
                      drawioRef.current.load({ xml: finalXml });
                    } catch (e) {
                      console.error('Failed to load diagram:', e);
                    }
                  }

                  setSessions(prev => prev.map(session => {
                    if (session.id === currentSessionId) {
                      return {
                        ...session,
                        drawIoXml: finalXml,
                        lastModified: Date.now()
                      };
                    }
                    return session;
                  }));
              }
              
              break;
            }

            case 'raw_text': {
              if (chunk.content) {
                agentTextContent += chunk.content;
                accumulatedContent += chunk.content;
                setMessages(prev => prev.map(m => m.id === agentMsgId ? { ...m, content: accumulatedContent, steps: [...accumulatedSteps] } : m));
              }
              break;
            }

            case 'user': {
              // AI returns a text response (not a diagram)
              if (chunk.content) {
                  // Sometimes the backend might leak raw JSON payload strings if the filter didn't catch it
                  let displayContent = chunk.content;
                  const trimmedContent = chunk.content.trim();
                  
                  // If it looks like a raw JSON string containing type="user", parse it and extract the content
                  if (trimmedContent.startsWith('{') && trimmedContent.includes('"type"') && trimmedContent.includes('"user"')) {
                      try {
                          const parsed = JSON.parse(trimmedContent);
                          if (parsed.content) {
                              displayContent = parsed.content;
                          }
                      } catch (e) {
                          // Try regex extraction if JSON parse fails due to streaming fragmentation
                          const match = trimmedContent.match(/"content"\s*:\s*"([^"]+)"/);
                          if (match && match[1]) {
                              displayContent = match[1].replace(/\\n/g, '\n');
                          }
                      }
                  }

                  // We treat 'user' type responses during stream as intermediate conversational feedback
                  agentTextContent += displayContent;
                  
                  // Make sure we only append actual text to the chat bubble, and skip any raw XML artifacts 
                  const isRawXmlLeak = displayContent.trim().startsWith('<mxCell') || displayContent.trim().startsWith('<mxGraphModel');
                  
                  if (!isRawXmlLeak) {
                      // Prevent duplicate consecutive lines
                      const newLines = displayContent.split('\n');
                      const currentLines = accumulatedContent.split('\n');
                      const lastLine = currentLines[currentLines.length - 1]?.trim();
                      
                      let shouldAppend = true;
                      if (newLines.length === 1 && newLines[0].trim() === lastLine) {
                          shouldAppend = false;
                      }

                      if (shouldAppend) {
                          const padding = accumulatedContent && !accumulatedContent.endsWith('\n\n') ? '\n\n' : '';
                          accumulatedContent += padding + displayContent;
                      }
                  }
                  
                  setMessages(prev => prev.map(m => m.id === agentMsgId ? { ...m, content: accumulatedContent, steps: [...accumulatedSteps] } : m));
              }
              break;
            }

            case 'status': {
              // Intermediate status message
              if (chunk.content) {
                 const text = chunk.content.trim();
                 // Filter out markdown code block syntax and raw JSON syntax from reasoning display
                 if (text !== '}' && text !== '{' && text !== ']' && text !== '[' && 
                     !text.startsWith('```') && 
                     !text.startsWith('"type":') &&
                     !text.startsWith('"id":') &&
                     !text.startsWith('"xml":') &&
                     !text.startsWith('<mxCell') &&
                     !text.startsWith('<mxGraphModel') &&
                     !text.includes('"drawio_node"') &&
                     !text.includes('"drawio_edge"') &&
                     !text.includes('"drawio_done"')) {
                     
                     // Don't append if the text is exactly the same as the last line to avoid repetitive status spam
                     if (!accumulatedReasoning.endsWith(text + '\n')) {
                         accumulatedReasoning += chunk.content + '\n';
                     }
                     updateStep(phase, currentPhaseLabel, chunk.content);
                     setStreamProgress(text.substring(0, 50) + '...');
                     setMessages(prev => prev.map(m => m.id === agentMsgId ? { ...m, reasoning: accumulatedReasoning, steps: [...accumulatedSteps] } : m));
                 }
              }
              break;
            }

            case 'token': {
              // Real-time token output
              if (chunk.content) {
                  const text = chunk.content.trim();
                  // We can accumulate tokens into the current step to show dynamic progress
                  // To avoid huge XML rendering lag, we can filter out JSON syntax
                  if (text !== '}' && text !== '{' && text !== ']' && text !== '[' && 
                     !text.startsWith('```') && 
                     !text.startsWith('"type":') &&
                     !text.startsWith('"id":') &&
                     !text.startsWith('"xml":') &&
                     !text.startsWith('<mxCell') &&
                     !text.startsWith('<mxGraphModel') &&
                     !text.includes('"drawio_node"') &&
                     !text.includes('"drawio_edge"') &&
                     !text.includes('"drawio_done"')) {
                     
                      // Avoid repeating the exact same text token
                      const stepIndex = accumulatedSteps.findIndex(s => s.phase === phase);
                      if (stepIndex >= 0 && accumulatedSteps[stepIndex].content.endsWith(chunk.content + '\n')) {
                          // Skip duplicate token append
                      } else {
                          updateStep(phase, currentPhaseLabel, chunk.content);
                          setMessages(prev => prev.map(m => m.id === agentMsgId ? { ...m, steps: [...accumulatedSteps] } : m));
                      }
                  }
              }
              break;
            }

            case 'error': {
              accumulatedContent += (accumulatedContent ? '\n\n' : '') + `❌ ${chunk.content}`;
              setMessages(prev => prev.map(m => m.id === agentMsgId ? { ...m, content: accumulatedContent, steps: [...accumulatedSteps] } : m));
              break;
            }

            case 'done': {
              // Stream completed explicitly by backend
              accumulatedSteps.forEach(s => { s.status = 'done'; });
              if (nodeCount > 0 || receivedDrawioDone) {
                accumulatedContent += (accumulatedContent ? '\n\n' : '') + `✅ 图表绘制完成！共 ${nodeCount} 个节点，${edgeCount} 条连线。`;
                setMessages(prev => prev.map(m => m.id === agentMsgId ? { ...m, content: accumulatedContent, steps: [...accumulatedSteps] } : m));
              } else if (!receivedDrawioDone && !agentTextContent && !hasIncrementalContent) {
                accumulatedContent += (accumulatedContent ? '\n\n' : '') + `⚠️ 未收到有效响应，请重试。`;
                setMessages(prev => prev.map(m => m.id === agentMsgId ? { ...m, content: accumulatedContent, steps: [...accumulatedSteps] } : m));
              } else {
                setMessages(prev => prev.map(m => m.id === agentMsgId ? { ...m, steps: [...accumulatedSteps] } : m));
              }
              setStreamPhase('done');
              break;
            }
          }
        },
        // onError
        (error: Error) => {
          console.error('Stream error:', error);
          
          // Only show error message if we didn't receive any content and it's not an AbortError
          if (error.name !== 'AbortError' && !hasIncrementalContent && !agentTextContent && nodeCount === 0) {
              accumulatedContent += (accumulatedContent ? '\n\n' : '') + `❌ 连接异常: ${error.message}`;
              setMessages(prev => prev.map(m => m.id === agentMsgId ? { ...m, content: accumulatedContent, steps: m.steps?.map(s => ({...s, status: 'done'})) } : m));
          } else {
              // If we already had content, just mark steps as done gracefully
              setMessages(prev => prev.map(m => m.id === agentMsgId ? { ...m, steps: m.steps?.map(s => ({...s, status: 'done'})) } : m));
          }
          
          setIsSending(false);
          setStreamPhase('');
          setStreamProgress('');
        },
        // onComplete
        () => {
          setIsSending(false);
          setStreamPhase('');
          setStreamProgress('');
          
          // Fallback if drawio_done wasn't explicitly sent
          if (!receivedDrawioDone && nodeCount > 0) {
              setMessages(prev => prev.map(m => {
                  if (m.id === agentMsgId) {
                      const finalContent = (m.content ? m.content + '\n\n' : '') + `✅ 图表绘制完成！共 ${nodeCount} 个节点，${edgeCount} 条连线。`;
                      return { ...m, content: finalContent, steps: m.steps?.map(s => ({...s, status: 'done'})) };
                  }
                  return m;
              }));
          } else if (!receivedDrawioDone && !agentTextContent && !hasIncrementalContent) {
              setMessages(prev => prev.map(m => {
                  if (m.id === agentMsgId) {
                      const finalContent = (m.content ? m.content + '\n\n' : '') + `⚠️ 未收到有效响应，请重试。`;
                      return { ...m, content: finalContent, steps: m.steps?.map(s => ({...s, status: 'done'})) };
                  }
                  return m;
              }));
          } else {
              setMessages(prev => prev.map(m => {
                  if (m.id === agentMsgId) {
                      return { ...m, steps: m.steps?.map(s => ({...s, status: 'done'})) };
                  }
                  return m;
              }));
          }
        }
      );

      streamAbortRef.current = controller;

    } catch (error) {
      console.error('Chat error:', error);
      setMessages(prev => [...prev, {
        id: Date.now().toString(),
        role: 'agent',
        content: error instanceof Error ? `Error: ${error.message}` : '发送失败，请重试。',
        timestamp: Date.now()
      }]);
      setIsSending(false);
      setStreamPhase('');
      setStreamProgress('');
    }
  };

  const handleSendMessage = async () => {
    if (!inputValue.trim() || isSending) return;
    
    const content = inputValue;
    setInputValue('');
    // Reset textarea height
    const textarea = document.querySelector('textarea');
    if (textarea) textarea.style.height = '80px';

    setIsSending(true);

    if (useHistoryContext && drawioRef.current && isDrawIoReady) {
        isExportingForChatRef.current = true;
        pendingMessageRef.current = content;
        try {
            drawioRef.current.exportDiagram({
                 format: 'xml' as any
             });
        } catch (e) {
            console.error("Export failed", e);
            performSendMessage(content, content);
        }
    } else {
        performSendMessage(content, content);
    }
  };

  useEffect(() => {
    if (!lastExportedData) return;
    
    if (isExportingForChatRef.current) {
        isExportingForChatRef.current = false;
        const xml = lastExportedData.data;
        const content = pendingMessageRef.current;
        const apiContent = `[Context: Current Draw.io XML]\n\`\`\`xml\n${xml}\n\`\`\`\n\n${content}`;
        performSendMessage(content, apiContent);
        return;
    }
    
    // Autosave handling
    if (isAutosaveRef.current) {
        isAutosaveRef.current = false;
        const xml = lastExportedData.data;
        setSessions(prev => prev.map(s => {
            if (s.id === currentSessionId) {
                return { ...s, drawIoXml: xml };
            }
            return s;
        }));
        return;
    }

    // Manual Export
    setImgData(lastExportedData.data);
  }, [lastExportedData]);

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSendMessage();
    }
  };

  const quickActions = [
    { label: 'UML类图', text: '请帮我绘制一个电商系统的UML类图，包含用户、订单、商品等类。' },
    { label: '时序图', text: '请帮我绘制一个用户登录注册的时序图，包含客户端、网关、认证服务、数据库。' },
    { label: '架构图', text: '请帮我绘制一个微服务系统架构图，包含接入层、业务逻辑层和数据层。' },
    { label: '业务流程图', text: '请帮我绘制一个电商购物流程图，包含商品浏览、加入购物车、下单、支付、发货等环节。' }
  ];

  return (
    <div className="flex flex-col h-screen w-full overflow-hidden bg-slate-50 text-slate-900 font-sans">
      {/* Header - Minimal & Clean */}
      <div className="h-14 px-6 bg-white border-b border-slate-100/60 flex items-center justify-between shrink-0 z-40 relative">
        <div className="flex items-center gap-3">
          <div className="bg-indigo-600 p-1.5 rounded-lg shadow-sm shadow-indigo-200">
            <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="white" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
               <rect x="3" y="3" width="18" height="18" rx="2" ry="2"></rect>
               <circle cx="8.5" cy="8.5" r="1.5"></circle>
               <polyline points="21 15 16 10 5 21"></polyline>
            </svg>
          </div>
          <h1 className="text-lg font-bold text-slate-800 tracking-tight">ai + draw.io</h1>
        </div>
        
        <div className="flex items-center gap-3">
          <button 
            onClick={() => router.push('/')}
            className="flex items-center gap-2 px-4 py-1.5 bg-white border border-slate-200 text-slate-600 rounded-lg hover:bg-slate-50 hover:border-slate-300 hover:text-slate-900 transition-all text-sm font-medium shadow-sm active:scale-95"
          >
            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="15 18 9 12 15 6"></polyline></svg>
            返回工作台
          </button>

          <a
            href="https://bugstack.cn/md/project/ai-agent-scaffold/ai-agent-scaffold.html"
            target="_blank"
            rel="noopener noreferrer"
            className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-500 hover:text-slate-800 bg-white border border-slate-200 hover:border-slate-300 rounded-lg transition-all shadow-sm"
          >
            <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M2 3h6a4 4 0 0 1 4 4v14a3 3 0 0 0-3-3H2z"/><path d="M22 3h-6a4 4 0 0 0-4 4v14a3 3 0 0 1 3-3h7z"/></svg>
            课程&源码
          </a>

          <div className="h-6 w-px bg-slate-200 mx-1"></div>

          <div className="flex items-center gap-2 px-3 py-1.5 bg-slate-50 rounded-full border border-slate-200 shadow-sm">
             <div className="w-2 h-2 rounded-full bg-green-500 shadow-[0_0_8px_rgba(34,197,94,0.4)]"></div>
             <span className="text-xs font-semibold text-slate-600">{currentUser || 'Guest'}</span>
          </div>

          <div className="h-6 w-px bg-slate-200 mx-1"></div>

          <button 
            onClick={exportDiagram}
            className="flex items-center gap-2 px-4 py-1.5 bg-white border border-slate-200 text-slate-600 rounded-lg hover:bg-slate-50 hover:border-slate-300 hover:text-slate-900 transition-all text-sm font-medium shadow-sm active:scale-95"
          >
            <Icons.Download className="w-4 h-4" />
            Export
          </button>

          <button
             onClick={handleLogout}
             className="p-2 text-slate-400 hover:text-red-600 hover:bg-red-50 rounded-lg transition-colors"
             title="Logout"
          >
            <Icons.Logout />
          </button>

          {!isChatOpen && (
            <button
              onClick={() => setIsChatOpen(true)}
              className="p-2 text-indigo-600 bg-indigo-50 hover:bg-indigo-100 rounded-lg transition-colors border border-indigo-100"
              title="Open Assistant"
            >
              <Icons.Chat />
            </button>
          )}
        </div>
      </div>
      
      {/* Main Layout */}
      <div className="flex flex-1 w-full overflow-hidden relative">
        {/* Sessions Sidebar */}
        <div className="w-64 bg-white text-slate-600 flex flex-col border-r border-slate-100/60 shrink-0 z-30">
          <div className="h-14 px-4 flex items-center justify-between border-b border-slate-100 shrink-0">
             <span className="font-semibold text-slate-800 flex items-center gap-2">
                <Icons.MessageSquare className="w-4 h-4 text-indigo-600" />
                绘图记录
             </span>
             <button 
                onClick={handleNewChat}
                className="p-1.5 text-slate-400 hover:text-indigo-600 hover:bg-indigo-50 rounded-md transition-all"
                title="New Chat"
             >
                <Icons.Plus className="w-5 h-5" />
             </button>
          </div>
          <div className="flex-1 overflow-y-auto p-2 space-y-1 scrollbar-thin scrollbar-thumb-slate-200 scrollbar-track-transparent">
             {[...sessions].sort((a, b) => b.lastModified - a.lastModified).map(session => (
                <div 
                  key={session.id}
                  onClick={() => handleSwitchSession(session.id)}
                  onDoubleClick={(e) => { e.stopPropagation(); handleDoubleClickSession(session); }}
                  className={`
                    group flex items-center gap-3 px-3 py-3 rounded-lg cursor-pointer transition-all border border-transparent
                    ${currentSessionId === session.id 
                      ? 'bg-indigo-50 text-indigo-700 border-indigo-100 shadow-sm' 
                      : 'hover:bg-slate-50 text-slate-600 hover:text-slate-900'
                    }
                  `}
                >
                  <div className="flex-1 min-w-0">
                     <div className={`text-sm font-medium truncate ${currentSessionId === session.id ? 'text-indigo-700' : 'text-slate-700 group-hover:text-slate-900'}`}>
                        {session.title}
                     </div>
                     <div className={`text-[10px] mt-0.5 ${currentSessionId === session.id ? 'text-indigo-400' : 'text-slate-400'}`}>
                        {new Date(session.lastModified).toLocaleDateString()} {new Date(session.lastModified).toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'})}
                     </div>
                  </div>
                  <button
                    onClick={(e) => handleDeleteSession(e, session.id)}
                    className={`
                      p-1.5 rounded-md transition-all opacity-0 group-hover:opacity-100
                      ${currentSessionId === session.id 
                        ? 'hover:bg-indigo-100 text-indigo-400 hover:text-indigo-700' 
                        : 'hover:bg-red-50 text-slate-400 hover:text-red-500'
                      }
                    `}
                    title="Delete"
                  >
                    <Icons.Trash className="w-4 h-4" />
                  </button>
                </div>
             ))}
             {sessions.length === 0 && (
                <div className="text-center py-10 text-xs text-slate-400">
                    No history yet
                </div>
             )}
          </div>
        </div>

        {/* Draw.io Canvas Area */}
        <div className="flex-1 relative bg-slate-50 h-full flex flex-col">
          <div className="flex-1 m-3 rounded-2xl overflow-hidden border border-slate-200 shadow-sm bg-white ring-1 ring-slate-100">
            <DrawIoEmbed 
              ref={drawioRef}
              autosave={true}
              onAutoSave={(data) => {
                if (currentSessionId && isDrawIoReady && !isExportingForChatRef.current) {
                   // Prefer using the XML directly from the autosave event if available
                   if (data && typeof data === 'object' && 'xml' in data) {
                       const xmlContent = (data as any).xml;
                       setSessions(prev => prev.map(s => {
                           if (s.id === currentSessionId) {
                               return { ...s, drawIoXml: xmlContent };
                           }
                           return s;
                       }));
                   } else {
                       // Fallback to export if no XML provided in event
                        isAutosaveRef.current = true;
                        drawioRef.current?.exportDiagram({ format: 'xml' as any });
                    }
                }
              }}
              onLoad={() => setIsDrawIoReady(true)}
              onExport={(data) => setLastExportedData({ data: data.data, timestamp: Date.now() })}
              urlParameters={{
                ui: 'atlas', // More modern UI theme for draw.io
                spin: true,
                libraries: true,
                saveAndExit: false,
                noSaveBtn: true,
                noExitBtn: true
              }}
            />
          </div>
        </div>

        {/* Chat Sidebar - Modern & Elegant */}
        <div 
          className={`
            border-l border-slate-100/60 bg-white flex flex-col transition-all duration-300 ease-[cubic-bezier(0.25,0.1,0.25,1)]
            ${isChatOpen ? 'w-[380px] translate-x-0' : 'w-0 translate-x-full opacity-0 overflow-hidden'}
            shadow-xl z-20
          `}
        >
          {/* Chat Header */}
          <div className="h-14 px-5 border-b border-slate-100 flex items-center justify-between shrink-0 bg-white/80 backdrop-blur-sm sticky top-0 z-10">
            <div className="flex items-center gap-3 flex-1 min-w-0">
              <div className="flex items-center justify-center w-8 h-8 rounded-full bg-gradient-to-br from-indigo-500 to-purple-600 text-white shadow-md shadow-indigo-200 shrink-0 ring-2 ring-white">
                <Icons.Sparkles className="w-4 h-4" />
              </div>
              <div className="flex-1 min-w-0">
                <select 
                  value={selectedAgentId} 
                  onChange={handleAgentChange}
                  className="w-full bg-transparent text-sm font-bold text-slate-800 focus:outline-none cursor-pointer truncate appearance-none pr-4"
                  style={{ backgroundImage: 'none' }}
                >
                  {agents.length === 0 && <option value="">Loading agents...</option>}
                  {agents.map(agent => (
                    <option key={agent.agentId} value={agent.agentId}>
                      {agent.agentName}
                    </option>
                  ))}
                </select>
                <div className="flex items-center gap-1.5 mt-0.5">
                   <span className="w-1.5 h-1.5 rounded-full bg-green-500 animate-pulse"></span>
                   <span className="text-[10px] text-slate-500 font-medium leading-tight">AI Assistant Online</span>
                </div>
              </div>
            </div>
            <div className="flex items-center gap-1">
                <button 
                  onClick={() => setIsChatOpen(false)}
                  className="p-1.5 text-slate-400 hover:text-slate-600 hover:bg-slate-100 rounded-md transition-all shrink-0"
                >
                  <Icons.Close className="w-5 h-5" />
                </button>
            </div>
          </div>

          {/* Messages Area */}
          <div className="flex-1 overflow-y-auto p-5 space-y-6 bg-slate-50/50 scrollbar-thin scrollbar-thumb-slate-200 scrollbar-track-transparent">
            {messages.map((msg, index) => {
              return (
                <div 
                  key={`${msg.id}-${index}`} 
                  className={`flex gap-3 ${msg.role === 'user' ? 'flex-row-reverse' : 'flex-row'}`}
                >
                  <div className={`
                    shrink-0 w-8 h-8 rounded-full flex items-center justify-center shadow-sm mt-1 ring-2 ring-white
                    ${msg.role === 'user' 
                      ? 'bg-indigo-100 text-indigo-600' 
                      : 'bg-white text-indigo-500 border border-slate-100'
                    }
                  `}>
                    {msg.role === 'user' ? <Icons.User className="w-5 h-5" /> : <Icons.Bot className="w-5 h-5" />}
                  </div>
                  
                  <div className="flex flex-col max-w-[85%] w-full">
                      <span className={`text-[10px] mb-1.5 font-medium ${msg.role === 'user' ? 'text-right text-slate-400' : 'text-left text-slate-400'}`}>
                          {msg.role === 'user' ? 'You' : 'Agent'}
                      </span>
                      
                      <div className={`flex flex-col gap-2 ${msg.role === 'user' ? 'items-end' : 'items-start'}`}>
                        {/* Steps / Reasoning Block */}
                        {msg.role === 'agent' && ((msg.steps && msg.steps.length > 0) || msg.reasoning) && (
                          <div className="w-full max-w-full">
                            <details className="w-full group/details open:pb-2" open={index === messages.length - 1 && isSending}>
                              <summary className="inline-flex items-center gap-2 cursor-pointer text-xs text-slate-500 hover:text-slate-700 font-medium select-none bg-white border border-slate-200 px-3 py-1.5 rounded-lg shadow-sm transition-all hover:border-slate-300">
                                 <Icons.Sparkles className="w-3.5 h-3.5 text-indigo-400" />
                                 <span className="group-open/details:hidden">展开执行步骤</span>
                                 <span className="hidden group-open/details:inline">收起执行步骤</span>
                              </summary>
                              <div className="mt-2 flex flex-col gap-2 p-3 bg-slate-50/50 border border-slate-200 rounded-xl shadow-sm text-sm text-slate-600 max-w-none overflow-x-auto">
                                 {msg.steps && msg.steps.length > 0 ? (
                                   msg.steps.map((step, idx) => (
                                     <div key={idx} className="flex flex-col gap-1.5 p-2 bg-white rounded-lg border border-slate-100 shadow-sm">
                                         <div className="flex items-center gap-2 font-medium text-slate-700">
                                             {step.status === 'running' ? (
                                                <Icons.Loader className="w-3.5 h-3.5 text-indigo-500" />
                                             ) : (
                                                <span className="text-green-500">✓</span>
                                             )}
                                             <span>{step.label}</span>
                                         </div>
                                         {step.content && (
                                             <div className="text-xs text-slate-500 pl-6 border-l-2 border-slate-100 ml-1.5 prose prose-sm prose-slate max-w-none prose-p:my-1 prose-pre:my-2 prose-pre:bg-slate-100 prose-pre:text-slate-700">
                                               <ReactMarkdown remarkPlugins={[remarkGfm]}>{step.content}</ReactMarkdown>
                                           </div>
                                         )}
                                     </div>
                                   ))
                                 ) : (
                                   <div className="p-2 bg-white rounded-lg border border-slate-100 shadow-sm prose prose-sm prose-slate max-w-none prose-p:my-1 prose-pre:my-2 prose-pre:bg-slate-100 prose-pre:text-slate-700">
                                     <ReactMarkdown remarkPlugins={[remarkGfm]}>{msg.reasoning || ''}</ReactMarkdown>
                                   </div>
                                 )}
                              </div>
                            </details>
                          </div>
                        )}

                        {/* Content Block */}
                        {msg.content && (
                          <div 
                            className={`
                                p-3.5 text-sm leading-relaxed shadow-sm whitespace-pre-wrap w-fit
                                ${msg.role === 'user' 
                                ? 'bg-indigo-600 text-white rounded-2xl rounded-tr-sm shadow-indigo-200' 
                                : 'bg-white border border-slate-200 text-slate-700 rounded-2xl rounded-tl-sm shadow-sm prose prose-sm prose-slate max-w-none overflow-x-auto prose-p:my-1 prose-pre:my-2 prose-pre:bg-slate-100 prose-pre:text-slate-700'
                                }
                            `}
                          >
                            {msg.role === 'user' ? (
                               msg.content
                            ) : (
                               <ReactMarkdown remarkPlugins={[remarkGfm]}>{msg.content}</ReactMarkdown>
                            )}
                          </div>
                        )}

                        {/* Empty state while generating */}
                        {index === messages.length - 1 && msg.role === 'agent' && !msg.content && !msg.reasoning && isSending && (
                           <div className="flex gap-1 items-center px-4 py-3 text-sm shadow-sm bg-white border border-indigo-100 text-indigo-600 rounded-2xl rounded-tl-sm">
                             <span className="w-1.5 h-1.5 rounded-full bg-indigo-500 animate-bounce" style={{ animationDelay: '0ms' }}></span>
                             <span className="w-1.5 h-1.5 rounded-full bg-indigo-500 animate-bounce" style={{ animationDelay: '150ms' }}></span>
                             <span className="w-1.5 h-1.5 rounded-full bg-indigo-500 animate-bounce" style={{ animationDelay: '300ms' }}></span>
                           </div>
                        )}
                      </div>
                  </div>
                </div>
              );
            })}

            <div ref={messagesEndRef} />
          </div>

          {/* Input Area */}
          <div className="p-4 bg-white border-t border-slate-100 shrink-0 relative z-20 shadow-[0_-4px_6px_-1px_rgba(0,0,0,0.02)]">
            {/* Quick Actions - Only show when chat is empty (just greeting) */}
            {messages.length <= 1 && (
              <div className="flex flex-wrap gap-2 mb-3 px-1 animate-in fade-in slide-in-from-bottom-2 duration-300">
                {quickActions.map((action, idx) => (
                  <button
                    key={idx}
                    onClick={() => {
                      setInputValue('');
                      performSendMessage(action.text, action.text);
                    }}
                    className="text-xs px-3 py-1.5 bg-indigo-50 text-indigo-600 rounded-full hover:bg-indigo-100 transition-colors border border-indigo-100 font-medium shadow-sm"
                  >
                    {action.label}
                  </button>
                ))}
              </div>
            )}

            {/* Canvas Context Toggle - above textarea */}
            <div className="flex items-center gap-2 mb-2 px-1">
                <button
                    onClick={() => setUseHistoryContext(!useHistoryContext)}
                    className={`
                        flex items-center gap-1 px-2 py-1 rounded-full text-[11px] font-medium transition-all border
                        ${useHistoryContext 
                            ? 'bg-indigo-50 text-indigo-600 border-indigo-200' 
                            : 'bg-white text-slate-400 border-slate-200 hover:text-slate-600'
                        }
                    `}
                >
                    <Icons.Layers className={`w-3 h-3 ${useHistoryContext ? 'text-indigo-500' : 'text-slate-400'}`} />
                    <span>画布上下文</span>
                </button>
            </div>

            <div className="relative flex items-end gap-2 bg-slate-50 p-2 rounded-2xl border border-slate-200 focus-within:border-indigo-400 focus-within:ring-4 focus-within:ring-indigo-100 focus-within:bg-white transition-all shadow-sm">
              <textarea
                value={inputValue}
                onChange={(e) => {
                  setInputValue(e.target.value);
                  e.target.style.height = 'auto';
                  e.target.style.height = Math.min(e.target.scrollHeight, 300) + 'px';
                }}
                onKeyDown={handleKeyDown}
                placeholder={isSending ? "AI 正在生成中..." : "输入您的问题，描述您的需求..."}
                disabled={isSending}
                className="flex-1 px-4 py-3 bg-transparent border-none focus:ring-0 text-[15px] text-slate-800 placeholder:text-slate-400 resize-none max-h-[300px] min-h-[80px] scrollbar-thin scrollbar-thumb-slate-200 scrollbar-track-transparent disabled:opacity-50 disabled:cursor-not-allowed leading-relaxed"
                rows={1}
                style={{ height: 'auto', minHeight: '80px' }}
              />
              <div className="flex gap-1 mb-0.5 shrink-0">
                  {isSending ? (
                    <button
                      onClick={handleStopStream}
                      className="p-2.5 rounded-lg transition-all duration-200 flex items-center justify-center bg-red-100 text-red-600 hover:bg-red-200 shadow-sm"
                      title="停止生成"
                    >
                      <Icons.Square className="w-4 h-4" />
                    </button>
                  ) : (
                    <button
                      onClick={handleSendMessage}
                      disabled={!inputValue.trim()}
                      className={`
                        p-2.5 rounded-lg transition-all duration-200 flex items-center justify-center
                        ${inputValue.trim()
                          ? 'bg-indigo-600 text-white shadow-md shadow-indigo-200 hover:bg-indigo-700 hover:scale-105 active:scale-95' 
                          : 'bg-slate-200 text-slate-400 cursor-not-allowed'
                        }
                      `}
                      title="发送消息"
                    >
                      <Icons.Send className="w-4 h-4" />
                    </button>
                  )}
                  <button
                    onClick={handleRestartSession}
                    disabled={isSending}
                    className="p-2.5 rounded-lg bg-white text-slate-400 hover:bg-slate-50 hover:text-indigo-600 transition-all duration-200 border border-slate-200 hover:border-indigo-100 shadow-sm disabled:opacity-50 disabled:cursor-not-allowed"
                    title="重启对话"
                  >
                    <Icons.Plus className="w-4 h-4" />
                  </button>
              </div>
            </div>

            {/* Model Selection - below textarea */}
            <div className="flex items-center gap-2 mt-2 px-1">
                <div className="relative flex items-center bg-white border border-slate-200 rounded-full shadow-sm hover:border-slate-300 transition-colors">
                    <Icons.Sparkles className={`w-3 h-3 ml-2 ${selectedCustomModelId !== 'default' ? 'text-indigo-500' : 'text-slate-400'}`} />
                    <select
                        value={selectedCustomModelId}
                        onChange={(e) => {
                            if (e.target.value === 'add_new') {
                                setShowApiConfig(true);
                                e.target.value = selectedCustomModelId;
                            } else {
                                setSelectedCustomModelId(e.target.value);
                                localStorage.setItem('ai_agent_selected_model', e.target.value);
                            }
                        }}
                        className="appearance-none bg-transparent border-none text-[11px] font-medium text-slate-600 focus:ring-0 py-1 pl-1 pr-5 cursor-pointer outline-none"
                    >
                        <option value="default">默认模型</option>
                        {customModels.filter(m => m.enabled).map(m => (
                            <option key={m.id} value={m.id}>{m.name || m.model}</option>
                        ))}
                        <option disabled>──────────</option>
                        <option value="add_new">+ 管理模型</option>
                    </select>
                    <div className="pointer-events-none absolute inset-y-0 right-0 flex items-center px-1.5 text-slate-400">
                        <svg className="fill-current h-3 w-3" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20"><path d="M9.293 12.95l.707.707L15.657 8l-1.414-1.414L10 10.828 5.757 6.586 4.343 8z"/></svg>
                    </div>
                </div>
                <span className="text-[10px] text-slate-400 ml-auto hidden sm:inline">
                    <kbd className="px-1 py-0.5 bg-slate-100 border border-slate-200 rounded text-slate-500">Enter</kbd> 发送
                </span>
            </div>
            <div className="text-center mt-1.5">
                <p className="text-[10px] text-slate-400">
                  {isSending ? (streamProgress || '生成中...') : ''}
                </p>
            </div>
          </div>
        </div>
      </div>

      {/* Export Modal - Polished */}
      {imgData && (
        <div className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center z-50 p-6 animate-in fade-in duration-200">
            <div className="bg-white p-0 rounded-2xl shadow-2xl max-h-[90vh] flex flex-col w-full max-w-4xl overflow-hidden animate-in zoom-in-95 duration-200 border border-white/20">
                <div className="flex justify-between items-center px-6 py-4 border-b border-slate-100 bg-slate-50/50">
                    <div className="flex items-center gap-3">
                        <div className="p-2 bg-green-100 text-green-600 rounded-lg">
                            <Icons.Download className="w-5 h-5" />
                        </div>
                        <div>
                            <h2 className="text-lg font-bold text-slate-800">Export Ready</h2>
                            <p className="text-xs text-slate-500">Your diagram has been successfully converted</p>
                        </div>
                    </div>
                    <button 
                        onClick={() => setImgData(null)}
                        className="p-2 text-slate-400 hover:text-slate-600 hover:bg-slate-100 rounded-full transition-colors"
                    >
                        <Icons.Close className="w-5 h-5" />
                    </button>
                </div>
                
                <div className="flex-1 overflow-auto bg-slate-50/50 p-8 flex items-center justify-center min-h-[400px]">
                    <div className="bg-white p-2 rounded shadow-sm border border-slate-200">
                        <img src={imgData} alt="Exported diagram" className="max-w-full h-auto object-contain" />
                    </div>
                </div>
                
                <div className="px-6 py-4 border-t border-slate-100 bg-white flex justify-end gap-3">
                   <button 
                        onClick={() => setImgData(null)}
                        className="px-5 py-2.5 text-slate-600 font-medium hover:bg-slate-100 rounded-lg transition-colors text-sm"
                    >
                        Close Preview
                    </button>
                    <a 
                        href={imgData} 
                        download="diagram.svg"
                        className="px-5 py-2.5 bg-indigo-600 text-white font-medium rounded-lg hover:bg-indigo-700 shadow-lg shadow-indigo-200 hover:shadow-indigo-300 transition-all text-sm flex items-center gap-2"
                    >
                        <Icons.Download className="w-4 h-4" />
                        Download File
                    </a>
                </div>
            </div>
        </div>
      )}

      {/* Rename Modal */}
      {isRenameModalOpen && (
        <div className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center z-50 p-6 animate-in fade-in duration-200">
            <div className="bg-white rounded-2xl shadow-2xl w-full max-w-md overflow-hidden animate-in zoom-in-95 duration-200 border border-white/20">
                <div className="px-6 py-4 border-b border-slate-100 bg-slate-50/50 flex justify-between items-center">
                    <h2 className="text-lg font-bold text-slate-800">Rename Session</h2>
                    <button 
                        onClick={() => setIsRenameModalOpen(false)}
                        className="p-1 text-slate-400 hover:text-slate-600 hover:bg-slate-100 rounded-full transition-colors"
                    >
                        <Icons.Close className="w-5 h-5" />
                    </button>
                </div>
                
                <div className="p-6">
                    <label className="block text-sm font-medium text-slate-700 mb-2">
                        Session Name
                    </label>
                    <input 
                        type="text" 
                        value={newSessionTitle}
                        onChange={(e) => setNewSessionTitle(e.target.value)}
                        onKeyDown={(e) => e.key === 'Enter' && handleRenameSave()}
                        className="w-full px-4 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all"
                        placeholder="Enter new name"
                        autoFocus
                    />
                </div>
                
                <div className="px-6 py-4 border-t border-slate-100 bg-slate-50/50 flex justify-end gap-3">
                   <button 
                        onClick={() => setIsRenameModalOpen(false)}
                        className="px-4 py-2 text-slate-600 font-medium hover:bg-slate-100 rounded-lg transition-colors text-sm"
                    >
                        Cancel
                    </button>
                    <button 
                        onClick={handleRenameSave}
                        className="px-4 py-2 bg-indigo-600 text-white font-medium rounded-lg hover:bg-indigo-700 shadow-lg shadow-indigo-200 hover:shadow-indigo-300 transition-all text-sm"
                    >
                        Save Changes
                    </button>
                </div>
            </div>
        </div>
      )}
      {/* Custom Models Settings Modal */}
      {showApiConfig && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center z-50 p-6 animate-in fade-in duration-200">
            <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl overflow-hidden animate-in zoom-in-95 duration-200 border border-white/20 flex flex-col max-h-[90vh]">
                <div className="px-6 py-4 border-b border-slate-100 bg-slate-50/50 flex justify-between items-center shrink-0">
                    <div className="flex items-center gap-2.5">
                      <div className="p-1.5 bg-indigo-50 rounded-lg">
                        <Icons.Sparkles className="w-4 h-4 text-indigo-500" />
                      </div>
                      <h2 className="text-base font-bold text-slate-800">自定义模型配置</h2>
                    </div>
                    <button
                        onClick={() => {
                          setShowApiConfig(false);
                          setEditingModel(null);
                        }}
                        className="p-1.5 text-slate-400 hover:text-slate-600 hover:bg-slate-100 rounded-lg transition-colors"
                    >
                        <Icons.Close className="w-4 h-4" />
                    </button>
                </div>
                
                <div className="flex flex-1 overflow-hidden">
                    {/* List of Models */}
                    <div className="w-1/3 border-r border-slate-100 bg-slate-50 flex flex-col">
                        <div className="p-3">
                            <button 
                                onClick={handleAddNewModel}
                                className="w-full flex items-center justify-center gap-2 py-2 bg-white border border-indigo-200 text-indigo-600 rounded-lg hover:bg-indigo-50 transition-colors shadow-sm text-sm font-medium"
                            >
                                <Icons.Plus className="w-4 h-4" /> 添加模型
                            </button>
                        </div>
                        <div className="flex-1 overflow-y-auto p-3 space-y-2">
                            {customModels.map(model => (
                                <div 
                                    key={model.id}
                                    onClick={() => setEditingModel(model)}
                                    className={`p-3 rounded-xl border cursor-pointer transition-all ${editingModel?.id === model.id ? 'bg-indigo-50 border-indigo-200 shadow-sm ring-1 ring-indigo-100' : 'bg-white border-slate-200 hover:border-indigo-100 hover:shadow-sm'}`}
                                >
                                    <div className="flex items-center justify-between mb-1">
                                        <div className="font-semibold text-sm text-slate-800 truncate pr-2">{model.name}</div>
                                        <div className="flex items-center gap-1 shrink-0">
                                            {/* Toggle Switch */}
                                            <button 
                                                onClick={(e) => {
                                                    e.stopPropagation();
                                                    const newModels = customModels.map(m => m.id === model.id ? {...m, enabled: !m.enabled} : m);
                                                    saveCustomModels(newModels);
                                                    if (!(!model.enabled) && selectedCustomModelId === model.id) {
                                                        setSelectedCustomModelId('default');
                                                    }
                                                }}
                                                className={`relative inline-flex h-4 w-7 items-center rounded-full transition-colors focus:outline-none ${model.enabled ? 'bg-indigo-500' : 'bg-slate-300'}`}
                                            >
                                                <span className={`inline-block h-3 w-3 transform rounded-full bg-white transition-transform ${model.enabled ? 'translate-x-3.5' : 'translate-x-0.5'}`} />
                                            </button>
                                            <button onClick={(e) => { e.stopPropagation(); handleDeleteModel(model.id); }} className="text-slate-400 hover:text-red-500 ml-1">
                                                <Icons.Trash className="w-3.5 h-3.5" />
                                            </button>
                                        </div>
                                    </div>
                                    <div className="text-[10px] text-slate-500 truncate">{model.model}</div>
                                </div>
                            ))}
                            {customModels.length === 0 && (
                                <div className="text-center text-xs text-slate-400 py-6">
                                    暂无自定义模型<br/>点击上方按钮添加
                                </div>
                            )}
                        </div>
                    </div>

                    {/* Edit Form */}
                    <div className="flex-1 p-6 overflow-y-auto bg-white">
                        {editingModel ? (
                            <div className="space-y-4 animate-in fade-in duration-200">
                                <div>
                                    <label className="block text-xs font-medium text-slate-700 mb-1">展示名称</label>
                                    <input type="text" value={editingModel.name} onChange={e => setEditingModel({...editingModel, name: e.target.value})} className="w-full px-3 py-2 text-sm border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all" placeholder="例如：我的GPT-4o" />
                                </div>
                                <div>
                                    <label className="block text-xs font-medium text-slate-700 mb-1">模型名称 (Model)</label>
                                    <input type="text" value={editingModel.model} onChange={e => setEditingModel({...editingModel, model: e.target.value})} className="w-full px-3 py-2 text-sm border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all" placeholder="例如：gpt-4o" />
                                </div>
                                <div>
                                    <label className="block text-xs font-medium text-slate-700 mb-1">Base URL</label>
                                    <input type="text" value={editingModel.baseUrl} onChange={e => setEditingModel({...editingModel, baseUrl: e.target.value})} className="w-full px-3 py-2 text-sm border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all" placeholder="例如：https://api.openai.com" />
                                </div>
                                <div>
                                    <label className="block text-xs font-medium text-slate-700 mb-1">API Key</label>
                                    <input type="password" value={editingModel.apiKey} onChange={e => setEditingModel({...editingModel, apiKey: e.target.value})} className="w-full px-3 py-2 text-sm border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all" placeholder="sk-..." />
                                </div>
                                <div>
                                    <label className="block text-xs font-medium text-slate-700 mb-1">Completions Path (可选)</label>
                                    <input type="text" value={editingModel.completionsPath} onChange={e => setEditingModel({...editingModel, completionsPath: e.target.value})} className="w-full px-3 py-2 text-sm border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all" placeholder="默认为 v1/chat/completions" />
                                </div>
                                <div className="pt-2 flex justify-end">
                                    <button onClick={handleSaveEditingModel} className="px-6 py-2 bg-indigo-600 text-white font-medium rounded-lg hover:bg-indigo-700 shadow-sm transition-all text-sm">
                                        保存配置
                                    </button>
                                </div>
                            </div>
                        ) : (
                            <div className="h-full flex flex-col items-center justify-center text-slate-400">
                                <Icons.Sparkles className="w-12 h-12 mb-3 opacity-20" />
                                <p className="text-sm">选择左侧模型进行编辑，或点击添加</p>
                            </div>
                        )}
                    </div>
                </div>
            </div>
        </div>
      )}
    </div>
  );
}
