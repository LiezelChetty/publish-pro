import type { PublishingSettings } from "./types";

const PRESET_KEY = "publish-pro-publishing-presets-v1";

export type PublishingPreset = {
  id: string;
  name: string;
  type: "pageNumbers" | "headerFooter" | "watermark";
  createdAt: string;
  settings: Partial<PublishingSettings>;
};

export function loadPublishingPresets(): PublishingPreset[] {
  try {
    const raw = localStorage.getItem(PRESET_KEY);
    return raw ? (JSON.parse(raw) as PublishingPreset[]) : [];
  } catch {
    return [];
  }
}

export function savePublishingPreset(preset: PublishingPreset) {
  const existing = loadPublishingPresets().filter((item) => item.id !== preset.id);
  localStorage.setItem(PRESET_KEY, JSON.stringify([preset, ...existing].slice(0, 24)));
}

export function deletePublishingPreset(id: string) {
  localStorage.setItem(PRESET_KEY, JSON.stringify(loadPublishingPresets().filter((item) => item.id !== id)));
}
