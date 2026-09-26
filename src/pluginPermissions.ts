import {PluginManager} from 'sn-plugin-lib';

const pending = new Map<string, Promise<boolean>>();

async function ensurePermission(permission: string, description: string): Promise<boolean> {
  const existing = pending.get(permission);
  if (existing) return existing;

  const request = (async () => {
    try {
      if (Number(await PluginManager.hasPermission(permission)) > 0) return true;
      return Number(await PluginManager.requestPermission(permission, description)) > 0;
    } catch {
      return false;
    }
  })();

  pending.set(permission, request);
  try {
    return await request;
  } finally {
    pending.delete(permission);
  }
}

export function ensureFileReadPermission(): Promise<boolean> {
  return ensurePermission(
    'plugin.permission.FILE:READ',
    'Allow DOCX to open the Word documents you pick.',
  );
}

export function ensureFileWritePermission(): Promise<boolean> {
  return ensurePermission(
    'plugin.permission.FILE:WRITE',
    'Allow DOCX to keep its log in EXPORT and, later, to save your edited copies.',
  );
}
