import { useEffect, useRef } from "react";
import { indentWithTab } from "@codemirror/commands";
import { python } from "@codemirror/lang-python";
import { indentUnit } from "@codemirror/language";
import { Compartment } from "@codemirror/state";
import { oneDark } from "@codemirror/theme-one-dark";
import { EditorView, keymap } from "@codemirror/view";
import { basicSetup } from "codemirror";

const MONO = '"SF Mono", Menlo, monospace';
const LOOK = EditorView.theme({
  "&": { height: "100%", fontSize: "12px" },
  ".cm-scroller": { fontFamily: MONO, overflow: "auto" },
});

// CodeMirror works in \n, so a model saved with CRLF only changes once someone edits it.
const normalize = (text: string) => text.replace(/\r\n?/g, "\n");

export function CodeEditor(props: {
  value: string;
  onChange: (value: string) => void;
  dark: boolean;
  /** Select this 1-based line, then call onJumped. */
  jump?: number | null;
  onJumped?: () => void;
}) {
  const host = useRef<HTMLDivElement>(null);
  const view = useRef<EditorView>(null);
  const theme = useRef(new Compartment());
  const latest = useRef(props);
  latest.current = props;

  useEffect(() => {
    const editor = new EditorView({
      doc: latest.current.value,
      parent: host.current!,
      extensions: [
        basicSetup,
        keymap.of([indentWithTab]),
        python(),
        indentUnit.of("    "),
        LOOK,
        theme.current.of(latest.current.dark ? oneDark : []),
        EditorView.contentAttributes.of({ "aria-label": "Model code" }),
        EditorView.updateListener.of((update) => {
          if (!update.docChanged) return;
          const text = update.state.doc.toString();
          if (text !== normalize(latest.current.value)) latest.current.onChange(text);
        }),
      ],
    });
    view.current = editor;
    return () => editor.destroy();
  }, []);

  useEffect(() => {
    const editor = view.current!;
    const text = normalize(props.value);
    if (editor.state.doc.toString() !== text) {
      editor.dispatch({ changes: { from: 0, to: editor.state.doc.length, insert: text } });
    }
  }, [props.value]);

  useEffect(() => {
    view.current!.dispatch({ effects: theme.current.reconfigure(props.dark ? oneDark : []) });
  }, [props.dark]);

  useEffect(() => {
    if (props.jump == null) return;
    const editor = view.current!;
    const { doc } = editor.state;
    const line = doc.line(Math.min(Math.max(props.jump, 1), doc.lines));
    editor.dispatch({
      selection: { anchor: line.from, head: line.to },
      effects: EditorView.scrollIntoView(line.from, { y: "center" }),
    });
    editor.focus();
    latest.current.onJumped?.();
  }, [props.jump]);

  return (
    <div
      ref={host}
      style={{
        flex: 1,
        minHeight: 240,
        overflow: "hidden",
        border: "1px solid var(--colorNeutralStroke1)",
        borderRadius: 4,
      }}
    />
  );
}
