// The page's touch layer (native DocxTouchLayer): an ordinary view that also tells the pen
// from a finger, so the pen selects and a finger turns pages. A plain View where the
// native part is missing (a JS-only test run): every touch is then the pen, as before.

import {UIManager, View, requireNativeComponent, type ViewProps} from 'react-native';

export const TouchLayer: React.ComponentType<ViewProps> = UIManager.getViewManagerConfig?.('DocxTouchLayer')
  ? requireNativeComponent<ViewProps>('DocxTouchLayer')
  : View;
