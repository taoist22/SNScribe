import React, {useEffect, useState} from 'react';
import {View} from 'react-native';
import {Reader} from './src/Reader';
import {QuoteCapture} from './src/QuoteCapture';
import {BUTTON_ID_QUOTE, getLastButtonEvent, subscribeToButtonEvents} from './src/pluginRouter';

/**
 * The editor always stays mounted (an open document and its unsaved edits survive); the
 * "Quote → DOCX" selection-toolbar button shows the capture screen over it.
 */
export default function App(): React.JSX.Element {
  const [capture, setCapture] = useState(getLastButtonEvent()?.id === BUTTON_ID_QUOTE ? 1 : 0);
  useEffect(
    () =>
      subscribeToButtonEvents(e => {
        // A new press reads a new selection: a new key remounts the capture screen.
        setCapture(c => (e.id === BUTTON_ID_QUOTE ? Math.abs(c) + 1 : 0));
      }),
    [],
  );
  return (
    <View style={{flex: 1}}>
      <Reader />
      {capture > 0 ? <QuoteCapture key={capture} onDone={() => setCapture(0)} /> : null}
    </View>
  );
}
