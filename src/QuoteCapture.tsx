import React, {useEffect, useState} from 'react';
import {Pressable, ScrollView, StyleSheet, Text, TextInput, View} from 'react-native';
import {PluginCommAPI, PluginDocAPI, PluginManager} from 'sn-plugin-lib';
import {QUOTES_KEY, fileName, isEpub, parseList, type SavedQuote} from './domain/quotes';
import {cleanQuote, findDoi} from './domain/reference';
import {Docx} from './services/native';

/**
 * "Quote → SNScribe": opened from the text-selection toolbar of a PDF or EPUB. Shows the
 * selected text (line-break hyphens joined, editable), its file and page, and a DOI found on
 * the first page; Keep adds it to the quote list DOCX's Edit ▸ Insert quote… offers.
 * (The selection toolbar keeps the selection alive, so getLastSelectedText works — as proven
 * in sn-commonplacer; PluginDocAPI is the only safe caller of it.)
 */

type Res<T> = {success?: boolean; result?: T} | null | undefined;

type Captured = {text: string; path: string; page: number; doi?: string};

async function capture(): Promise<Captured | string> {
  const pathRes = (await PluginCommAPI.getCurrentFilePath()) as Res<string>;
  const path = pathRes?.success && pathRes.result ? pathRes.result : '';
  const textRes = (await PluginDocAPI.getLastSelectedText()) as Res<string>;
  const text = textRes?.success && textRes.result ? cleanQuote(textRes.result) : '';
  if (!text) {
    return 'No text is selected. Select a passage, then tap Quote → SNScribe.';
  }
  const pageRes = (await PluginCommAPI.getCurrentPageNum()) as Res<number>;
  const page = pageRes?.success && typeof pageRes.result === 'number' ? pageRes.result + 1 : 1;
  // A DOI is printed on an article's first page (or the current one); EPUBs carry their own details.
  let doi: string | undefined;
  if (!isEpub(path)) {
    for (const p of [0, page - 1]) {
      try {
        const t = (await PluginDocAPI.getCurrentDocText(p)) as Res<string>;
        doi = t?.success && t.result ? findDoi(t.result) : undefined;
      } catch {
        doi = undefined;
      }
      if (doi) {
        break;
      }
    }
  }
  return {text, path, page, ...(doi ? {doi} : {})};
}

export function QuoteCapture({onDone}: {onDone: () => void}): React.JSX.Element {
  const [got, setGot] = useState<Captured | null>(null);
  const [text, setText] = useState('');
  const [page, setPage] = useState('');
  const [message, setMessage] = useState('Reading the selection…');
  const [kept, setKept] = useState(false);

  useEffect(() => {
    let live = true;
    capture()
      .then(r => {
        if (!live) {
          return;
        }
        if (typeof r === 'string') {
          setMessage(r);
          return;
        }
        setGot(r);
        setText(r.text);
        setPage(String(r.page));
        setMessage('');
      })
      .catch(e => live && setMessage(`Could not read the selection: ${e instanceof Error ? e.message : String(e)}`));
    return () => {
      live = false;
    };
  }, []);

  const close = () => {
    onDone();
    PluginManager.closePluginView();
  };

  const keep = async () => {
    if (!got || !text.trim()) {
      return;
    }
    try {
      const list = parseList<SavedQuote>(await Docx?.load(QUOTES_KEY));
      const q: SavedQuote = {
        id: Date.now().toString(36),
        text: text.trim(),
        path: got.path,
        page: Math.max(1, parseInt(page, 10) || got.page),
        ...(got.doi ? {doi: got.doi} : {}),
        captured: new Date().toISOString(),
      };
      await Docx?.store(QUOTES_KEY, JSON.stringify([q, ...list]));
      setKept(true);
      setMessage(`Kept (${list.length + 1} waiting). In SNScribe: Edit → Insert quote…`);
      setTimeout(close, 1500);
    } catch (e) {
      setMessage(`Not kept: ${e instanceof Error ? e.message : String(e)}`);
    }
  };

  return (
    <View style={styles.screen}>
      <ScrollView contentContainerStyle={styles.panel} keyboardShouldPersistTaps="always">
        <Text allowFontScaling={false} style={styles.title}>
          {'Quote → SNScribe'}
        </Text>
        {got ? (
          <>
            <TextInput style={styles.quote} value={text} onChangeText={setText} multiline allowFontScaling={false} />
            <Text allowFontScaling={false} style={styles.note}>
              {'Check the text: words broken across lines are joined, but a PDF can still carry odd hyphens or breaks.'}
            </Text>
            <Text allowFontScaling={false} style={styles.source} numberOfLines={2}>
              {fileName(got.path)}
            </Text>
            <View style={styles.row}>
              <Text allowFontScaling={false} style={styles.label}>
                {isEpub(got.path) ? 'Page (EPUB pages vary; APA may want a chapter instead)' : 'Page'}
              </Text>
              <TextInput style={styles.page} value={page} onChangeText={setPage} keyboardType="number-pad" allowFontScaling={false} />
            </View>
            {got.doi ? (
              <Text allowFontScaling={false} style={styles.note}>
                {`DOI found: ${got.doi}`}
              </Text>
            ) : null}
          </>
        ) : null}
        {message ? (
          <Text allowFontScaling={false} style={styles.message}>
            {message}
          </Text>
        ) : null}
        <View style={styles.row}>
          {got && !kept ? (
            <Pressable onPress={keep} style={styles.button}>
              <Text allowFontScaling={false} style={styles.buttonText}>
                {'Keep quote'}
              </Text>
            </Pressable>
          ) : null}
          <Pressable onPress={close} style={styles.button}>
            <Text allowFontScaling={false} style={styles.buttonText}>
              {kept ? 'Close' : 'Cancel'}
            </Text>
          </Pressable>
        </View>
      </ScrollView>
    </View>
  );
}

const styles = StyleSheet.create({
  screen: {...StyleSheet.absoluteFillObject, backgroundColor: '#fff'},
  panel: {padding: 24},
  title: {color: '#000', fontSize: 24, fontWeight: '700', marginBottom: 12},
  quote: {minHeight: 180, borderWidth: 1, borderColor: '#000', padding: 10, fontSize: 18, color: '#000', textAlignVertical: 'top'},
  note: {color: '#333', fontSize: 15, marginTop: 8},
  source: {color: '#000', fontSize: 18, fontWeight: '700', marginTop: 14},
  row: {flexDirection: 'row', alignItems: 'center', marginTop: 12, flexWrap: 'wrap'},
  label: {color: '#000', fontSize: 16, marginRight: 10, flexShrink: 1},
  page: {width: 90, height: 44, borderWidth: 1, borderColor: '#000', paddingHorizontal: 8, fontSize: 18, color: '#000'},
  message: {color: '#000', fontSize: 17, fontWeight: '700', marginTop: 14},
  button: {borderWidth: 1, borderColor: '#000', borderRadius: 6, paddingVertical: 10, paddingHorizontal: 18, marginRight: 12, marginTop: 8},
  buttonText: {color: '#000', fontSize: 18},
});
