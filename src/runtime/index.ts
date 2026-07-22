import { browserRuntime } from "./browser";
import { desktopRuntime, isTauriRuntime, runtimeFileToBrowserFile } from "./desktop";
import type { RuntimeAdapter, RuntimeFile } from "./types";

export const runtime: RuntimeAdapter = isTauriRuntime() ? desktopRuntime : browserRuntime;

export { runtimeFileToBrowserFile, isTauriRuntime };
export type { RuntimeFile };
