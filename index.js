import {AppRegistry, Image} from 'react-native';
import App from './App';
import {name as appName} from './app.json';
import {PluginManager} from 'sn-plugin-lib';
import {installPluginRouter} from './src/pluginRouter';

const BUTTON_TYPE_TOOLBAR = 1;
const TOOLBAR_BUTTON_ID = 100;
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
