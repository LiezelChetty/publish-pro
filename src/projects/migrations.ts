import { assertSupportedManifest, PROJECT_FORMAT_VERSION, type ProjectManifest } from "./schema";

export function migrateProjectManifest(value: unknown): ProjectManifest {
  assertSupportedManifest(value);

  if (value.version === PROJECT_FORMAT_VERSION) return value;

  throw new Error(`Project version ${value.version} cannot be migrated by this Publish Pro build.`);
}
