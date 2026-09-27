import {AppRegistry, Image} from 'react-native';
import App from './App';
import {name as appName} from './app.json';
import {PluginManager} from 'sn-plugin-lib';
import {installPluginRouter} from './src/pluginRouter';

const BUTTON_TYPE_TOOLBAR = 1;
const BUTTON_TYPE_TEXT_SELECTION = 3;
const TOOLBAR_BUTTON_ID = 100;
const QUOTE_BUTTON_ID = 101;
const SHOW_TYPE_WITH_UI = 1;

AppRegistry.registerComponent(appName, () => App);

PluginManager.init();
installPluginRouter();

// NOTE and DOC: a plugin can only be launched from a note or a PDF/EPUB toolbar.
PluginManager.registerButton(BUTTON_TYPE_TOOLBAR, ['NOTE', 'DOC'], {
  id: TOOLBAR_BUTTON_ID,
  name: 'DOCX',
  icon: Image.resolveAssetSource(require('./assets/docx.png')).uri,
  showType: SHOW_TYPE_WITH_UI,
});

// Type 3 = the text-selection toolbar of a PDF/EPUB (DOC only): the selection is still live
// when the capture screen opens, so getLastSelectedText returns it (proven in sn-commonplacer).
PluginManager.registerButton(BUTTON_TYPE_TEXT_SELECTION, ['DOC'], {
  id: QUOTE_BUTTON_ID,
  name: 'Quote → DOCX',
  icon: Image.resolveAssetSource(require('./assets/docx.png')).uri,
  showType: SHOW_TYPE_WITH_UI,
});
