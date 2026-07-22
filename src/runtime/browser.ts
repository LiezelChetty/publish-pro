import type { RuntimeAdapter } from "./types";

export const browserRuntime: RuntimeAdapter = {
  isDesktop: false,
  async openFiles() {
    return null;
  },
  async readPath() {
    return null;
  },
  async saveFile() {
    return null;
  },
  async setWindowTitle() {
    return;
  },
  async revealPath() {
    return;
  },
  async openPath() {
    return;
  },
  async forceClose() {
    return;
  },
  async loadAutosave() {
    return null;
  },
  async saveAutosave() {
    return null;
  },
  async clearAutosave() {
    return;
  },
  async loadRecentProjects() {
    return null;
  },
  async addRecentProject() {
    return;
  },
  async removeRecentProject() {
    return;
  },
  async clearRecentProjects() {
    return;
  },
  async listenMenuActions() {
    return () => undefined;
  },
  async listenLaunchFiles() {
    return () => undefined;
  },
};
