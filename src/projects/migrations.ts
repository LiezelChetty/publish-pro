import { createDefaultPublishingSettings, mergePublishingSettings } from "../publishing/defaults";
import { assertSupportedManifest, PROJECT_FORMAT_VERSION, type ProjectManifest } from "./schema";

export function migrateProjectManifest(value: unknown): ProjectManifest {
  if (!value || typeof value !== "object") {
    throw new Error("Project manifest is missing or invalid.");
  }
  const input = value as Omit<Partial<ProjectManifest>, "version" | "publishingSettings"> & { version?: number; publishingSettings?: unknown };
  const version = input.version;
  if (version === 1) {
    assertSupportedManifest({ ...input, publishingSettings: createDefaultPublishingSettings() });
    return {
      ...input,
      version: PROJECT_FORMAT_VERSION,
      publishingSettings: mergePublishingSettings(input.publishingSettings),
    } as ProjectManifest;
  }

  assertSupportedManifest(value);
  if (value.version === PROJECT_FORMAT_VERSION) {
    return {
      ...value,
      publishingSettings: mergePublishingSettings(value.publishingSettings),
    };
  }

  throw new Error(`Project version ${value.version} cannot be migrated by this Publish Pro build.`);
}
