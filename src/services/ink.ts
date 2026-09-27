// The firmware pen pad (DocxInkModule / DocxInkViewManager). Undocumented vendor internals:
// every call fails soft, and the feature is offered only when isInkAvailable() says so.

import {NativeModules, requireNativeComponent, type ViewStyle} from 'react-native';

type DocxInkNative = {
  isAvailable(): Promise<boolean>;
  activate(): Promise<string>;
  clear(): Promise<string>;
  deactivate(): Promise<string>;
  save(path: string): Promise<{path?: string; width?: number; height?: number; strokes?: number; empty?: boolean}>;
  notesDir(key: string): Promise<string>;
  remove(path: string): Promise<boolean>;
};

export const DocxInk: DocxInkNative | undefined = NativeModules.DocxInk;

/** The pen engine's host. FIXED size, never moved while the engine is live. */
export const InkSurfaceView = requireNativeComponent<{style?: ViewStyle}>('DocxInkSurface');

export async function isInkAvailable(): Promise<boolean> {
  try {
    return (await DocxInk?.isAvailable()) ?? false;
  } catch {
    return false;
  }
}

/** Binds the engine once the surface is laid out (retries while it is not). */
export async function activateInk(): Promise<string> {
  for (let i = 0; i < 20; i++) {
    let r = 'unavailable';
    try {
      r = (await DocxInk?.activate()) ?? 'unavailable';
    } catch {
      r = 'error';
    }
    if (r !== 'not-laid-out' && r !== 'no-surface') {
      return r;
    }
    await new Promise(res => setTimeout(res, 100));
  }
  return 'not-laid-out';
}

/** MUST run on every way out of the pad. */
export async function deactivateInk(): Promise<void> {
  try {
    await DocxInk?.deactivate();
  } catch {
    // Nothing to release.
  }
}
