/**
 * pluginRouter — single source of truth for plugin button press events.
 *
 * Installs exactly one listener with sn-plugin-lib and fans out to any
 * subscribers, so components can react to button events without each
 * registering their own listener (which the library disallows).
 *
 * App.tsx reads `getLastButtonEvent()` synchronously on first render to
 * pick the right initial view, then subscribes for any subsequent events.
 * sn-plugin-lib replays the cached `lastButtonEventMsg` to newly-registered
 * listeners within a 1-second window, which is enough to capture the initial
 * trigger into module state before components mount.
 */
import {PluginManager} from 'sn-plugin-lib';

export const BUTTON_ID_TOOLBAR = 100;

export type ButtonEvent = {
  pressEvent: number;
  id: number;
  name: string;
  color: number;
  icon: string;
  bgColor: number;
};

export type ButtonSubscriber = (event: ButtonEvent) => void;

let lastEvent: ButtonEvent | null = null;
const subscribers = new Set<ButtonSubscriber>();
let installed = false;

export function installPluginRouter(): void {
  if (installed) {
    return;
  }
  installed = true;
  PluginManager.registerButtonListener({
    onButtonPress(event: ButtonEvent) {
      console.log('[PLUGIN_ROUTER] onButtonPress', JSON.stringify(event));
      lastEvent = event;
      for (const fn of subscribers) {
        try {
          fn(event);
        } catch (e) {
          console.error('[PLUGIN_ROUTER] subscriber threw', e);
        }
      }
    },
  });
}

export function getLastButtonEvent(): ButtonEvent | null {
  return lastEvent;
}

export function subscribeToButtonEvents(fn: ButtonSubscriber): () => void {
  subscribers.add(fn);
  return () => {
    subscribers.delete(fn);
  };
}

export const __testing__ = {
  reset(): void {
    lastEvent = null;
    subscribers.clear();
    installed = false;
  },
  getSubscriberCount(): number {
    return subscribers.size;
  },
  isInstalled(): boolean {
    return installed;
  },
};
