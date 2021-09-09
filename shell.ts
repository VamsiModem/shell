// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
/* eslint-disable @typescript-eslint/no-explicit-any */

import { AxiosInstance } from 'axios';

import { IDiagnostic } from './diagnostic';
import type {
  DialogInfo,
  LuFile,
  LgFile,
  QnAFile,
  LuIntentSection,
  LgTemplate,
  DialogSchemaFile,
  LuProviderType,
} from './indexers';
import type { JSONSchema7, SDKKinds } from './schema';
import { Skill } from './indexers';
import type { ILUFeaturesConfig, SkillSetting, UserSettings, DialogSetting } from './settings';
import { MicrosoftIDialog } from './sdk';
import { FeatureFlagKey } from './featureFlags';
import { TelemetryClient } from './telemetry';
import { CurrentUser } from './auth';

/** Recursively marks all properties as optional. */
// type AllPartial<T> = {
//   [P in keyof T]?: T[P] extends (infer U)[] ? AllPartial<U>[] : T[P] extends object ? AllPartial<T[P]> : T[P];
// };

export type HttpClient = AxiosInstance;

export type ZoomInfo = {
  rateList: number[];
  maxRate: number;
  minRate: number;
  currentRate: number;
};

export type EditorSchema = {
  content?: {
    fieldTemplateOverrides: any;
    SDKOverrides?: any;
  };
};

type UISchema = {
  [key: string]: {
    form?: any;
    flow?: any;
    menu?: any;
  };
};

export type BotSchemas = {
  default?: JSONSchema7;
  sdk?: any;
  ui?: { content: UISchema };
  uiOverrides?: { content: UISchema };
  diagnostics?: any[];
};

// Only used in multi-bot scenarios
export type DisabledMenuActions = {
  kind: SDKKinds;
  reason: string;
};

// Notifications are only used in extensions
export type NotificationLink = { label: string; onClick: () => void };

// Notifications are only used in extensions
export type Notification = {
  type: 'info' | 'warning' | 'error' | 'pending' | 'success' | 'question' | 'congratulation' | 'custom';
  title: string;
  description?: string;
  retentionTime?: number;
  link?: NotificationLink;
  links?: NotificationLink[];
  leftLinks?: NotificationLink[];
  rightLinks?: NotificationLink[];
  icon?: string;
  color?: string;
  read?: boolean;
  hidden?: boolean;
  onRenderCardContent?: ((props: Notification) => JSX.Element) | React.FC<any>;
  data?: Record<string, any>;
  onDismiss?: (id: string) => void;
};

export type ApplicationContextApi = {
  navigateTo: (to: string, opts?: { state?: any; replace?: boolean }) => void; // Used in Lg and Luis ui-plugins
  updateUserSettings: (settings: Partial<UserSettings>) => void; // Used in Lg and Luis ui-plugins
  announce: (message: string) => void; // Used in flow
  addCoachMarkRef: (ref: { [key: string]: any }) => void; // Used in flow
  isFeatureEnabled: (featureFlagKey: FeatureFlagKey) => boolean; // Not used in form, flow, or ui-plugins
  setApplicationLevelError: (err: any) => void; // Not used in form, flow, or ui-plugins
  confirm: (title: string, subTitle: string, settings?: any) => Promise<boolean>; // Not used in form, flow, or ui-plugins
  updateFlowZoomRate: (currentRate: number) => void; // Used in flow
  toggleFlowComments: () => void; // Used in flow
  telemetryClient: TelemetryClient; // Used in form, Lg ui-plugin, and Luis ui-plugin; however, PVA uses a different telemetry architecture
  addNotification: (notification: Notification) => string; // Not used in form, flow, or ui-plugins
  deleteNotification: (id: string) => void; // Not used in form, flow, or ui-plugins
  markNotificationAsRead: (id: string) => void; // Not used in form, flow, or ui-plugins
  hideNotification: (id: string) => void; // Not used in form, flow, or ui-plugins
};

export type ApplicationContext = {
  locale: string; // Used in form, Lg ui-plugin, Luis ui-plugin, cross-trained ui-plugin, and orchestrator ui-plugin
  hosted: boolean; // Used in flow; however can probably be removed since PVA2 does not use electron
  userSettings: UserSettings; // Used in flow, Lg ui-plugin, and Luis ui-plugin
  skills: Record<string, Skill>; // Used in select skill ui-plugin
  skillsSettings: Record<string, SkillSetting>; // Used in select skill ui-plugin
  // TODO: remove
  schemas: BotSchemas; // Used in form, flow, Luis ui-plugin, cross-trained ui-plugin, and orchestrator ui-plugin
  flowZoomRate: ZoomInfo; // Used in flow
  flowCommentsVisible: boolean; // Used in flow

  httpClient: HttpClient; // Not used in form, flow, or ui-plugins
};

// Not used in form, flow, or ui-plugins
export type AuthContext = {
  currentUser: CurrentUser;
  currentTenant: string;
  isAuthenticated: boolean;
  showAuthDialog: boolean;
};

// Not used in form, flow, or ui-plugins
export type AuthContextApi = {
  requireUserLogin: (tenantId?: string, options?: { requireGraph: boolean }) => void;
};

export type LuContextApi = {
  getLuIntent: (id: string, intentName: string) => LuIntentSection | undefined; // Not used in flow, form, or ui-plugins
  getLuIntents: (id: string) => LuIntentSection[]; // Not used in flow, form, or ui-plugins
  addLuIntent: (id: string, intentName: string, intent: LuIntentSection) => Promise<LuFile[] | undefined>; // Not used in flow, form, or ui-plugins
  updateLuFile: (id: string, content: string) => Promise<void>; // Not used in flow, form, or ui-plugins
  updateLuIntent: (id: string, intentName: string, intent: LuIntentSection) => Promise<LuFile[] | undefined>; // Not used in flow, form, or ui-plugins
  debouncedUpdateLuIntent: (id: string, intentName: string, intent: LuIntentSection) => Promise<LuFile[] | undefined>; // Used in Luis ui-plugin
  renameLuIntent: (id: string, intentName: string, newIntentName: string) => Promise<LuFile[] | undefined>; // Used in Luis ui-plugin
  removeLuIntent: (id: string, intentName: string) => Promise<LuFile[] | undefined>; // Not used in flow, form, or ui-plugins
};

export type LgContextApi = {
  getLgTemplates: (id: string) => LgTemplate[]; // Not used in flow, form, or ui-plugins
  copyLgTemplate: (id: string, fromTemplateName: string, toTemplateName?: string) => Promise<LgFile[] | undefined>; // Not used in flow, form, or ui-plugins
  addLgTemplate: (id: string, templateName: string, templateStr: string) => Promise<LgFile[] | undefined>; // Not used in flow, form, or ui-plugins
  updateLgFile: (id: string, content: string) => Promise<void>; // Not used in flow, form, or ui-plugins
  updateLgTemplate: (id: string, templateName: string, templateStr: string) => Promise<LgFile[] | undefined>; // Not used in flow, form, or ui-plugins
  debouncedUpdateLgTemplate: (id: string, templateName: string, templateStr: string) => Promise<LgFile[] | undefined>; // Used in Lg ui-plugin
  removeLgTemplate: (id: string, templateName: string) => Promise<LgFile[] | undefined>; // Used in Lg ui-plugin
  removeLgTemplates: (id: string, templateNames: string[]) => Promise<LgFile[] | undefined>; // Not used in flow, form, or ui-plugins
};

export type ProjectContextApi = {
  getMemoryVariables: (projectId: string, options?: { signal: AbortSignal }) => Promise<string[]>; // Used in form and Lg ui-plugin
  getDialog: (dialogId: string) => any; // Used in flow
  saveDialog: (dialogId: string, newDialogData: any) => any; // Used in flow
  reloadProject: () => void; // Not used in form, flow, or ui-plugins
  navTo: (path: string) => void; // Used in select dialog ui-plugin - how is this different from navigateTo?

  updateQnaContent: (id: string, content: string) => void; // Not used in form, flow, or ui-plugins
  updateRegExIntent: (id: string, intentName: string, pattern: string) => void; // Used in form
  renameRegExIntent: (id: string, intentName: string, newIntentName: string) => void; // Used in composer ui-plugin
  updateIntentTrigger: (id: string, intentName: string, newIntentName: string) => void; // Used in Luis ui-plugin
  createDialog: (actions?: any[]) => Promise<string | null>; // Used in flow and select dialog ui-plugin
  commitChanges: () => void; // Used in Lg ui-plugin and Luis ui-plugin
  displayManifestModal: (manifestId: string) => void; // Used in select skill ui-plugin
  updateDialogSchema: (_: DialogSchemaFile) => Promise<void>; // Used in schema-editor ui-plugin
  createTrigger: (id: string, formData, autoSelected?: boolean) => void; // Not used in form, flow, or ui-plugins
  createQnATrigger: (id: string) => void; // Not used in form, flow, or ui-plugins
  stopBot: (projectId: string) => void; // Not used in form, flow, or ui-plugins and unnecessary for PVA 2.0
  updateSkill: (skillId: string, skillsData: { skill: Skill; selectedEndpointIndex: number }) => Promise<void>; // Used in  select skill dialog ui-plugin
  updateRecognizer: (projectId: string, dialogId: string, kind: LuProviderType) => void; // Used in flow, cross-trained ui-plugin, and orchestrator ui-plugin
};

// Not used in form, flow, or ui-plugins
export type BotInProject = {
  dialogs: DialogInfo[];
  projectId: string;
  name: string;
  isRemote: boolean;
  isRootBot: boolean;
  diagnostics: IDiagnostic[];
  error: { [key: string]: any };
  buildEssentials: { [key: string]: any };
  isPvaSchema: boolean;
  setting: DialogSetting;
};

export type ProjectContext = {
  botName: string;
  projectId: string;
  projectCollection: BotInProject[]; // Not used in form, flow, or ui-plugins
  dialogs: DialogInfo[];
  topics: DialogInfo[];
  dialogSchemas: DialogSchemaFile[];
  lgFiles: LgFile[];
  luFiles: LuFile[];
  luFeatures: ILUFeaturesConfig; // Used in Lu plugin
  qnaFiles: QnAFile[]; // Used in form, cross-trained ui-plugin, and orchestrator ui-plugin
  skills: Record<string, Skill>; // Select skill ui-plugin
  skillsSettings: Record<string, SkillSetting>; // Select skill ui-plugin
  schemas: BotSchemas; // Duplicate of schema in ApplicationContext
  forceDisabledActions: DisabledMenuActions[]; // Used in flow - related to internal skills which PVA does not support
  settings: DialogSetting; // Not used in form, flow, or ui-plugins
};

export type ActionContextApi = {
  constructAction: (dialogId: string, action: MicrosoftIDialog) => Promise<MicrosoftIDialog>; // Not used in form, flow, or ui-plugins
  constructActions: (dialogId: string, actions: MicrosoftIDialog[]) => Promise<MicrosoftIDialog[]>; // Used in flow
  copyAction: (dialogId: string, action: MicrosoftIDialog) => Promise<MicrosoftIDialog>; // Used in flow
  copyActions: (dialogId: string, actions: MicrosoftIDialog[]) => Promise<MicrosoftIDialog[]>; // Used in flow
  deleteAction: (dialogId: string, action: MicrosoftIDialog) => Promise<void>; // Used in flow
  deleteActions: (dialogId: string, actions: MicrosoftIDialog[]) => Promise<void>; // Used in flow
  actionsContainLuIntent: (action: MicrosoftIDialog[]) => boolean; // Used in flow
};

export type DialogEditingContextApi = {
  saveData: <T = any>(newData: T, updatePath?: string, callback?: () => void | Promise<void>) => Promise<void>; // Used in flow
  onOpenDialog: (dialogId: string) => Promise<void>; // Used in flow
  onFocusSteps: (stepIds: string[], focusedTab?: string) => Promise<void>; // Used in flow
  onFocusEvent: (eventId: string) => Promise<void>; // Used in flow
  onSelect: (ids: string[]) => void; // Used in flow
  onCopy: (clipboardActions: any[]) => void; // Used in flow
  undo: () => void; // Used in flow
  redo: () => void; // Used in flow
};

export type DialogEditingContext = {
  currentDialog: DialogInfo;
  designerId: string;
  dialogId: string;
  clipboardActions: any[];
  focusedEvent: string;
  focusedActions: string[];
  focusedSteps: string[];
  focusedTab?: string;
  focusPath: string;
};

export type ShellData = ApplicationContext & AuthContext & ProjectContext & DialogEditingContext;

export type ShellApi = ApplicationContextApi &
  AuthContextApi &
  ProjectContextApi &
  DialogEditingContextApi &
  LgContextApi &
  LuContextApi &
  ActionContextApi;

export type Shell = {
  api: ShellApi;
  data: ShellData;
};
