import { useEffect } from "react";
import {
  Button,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  Text,
  Toast,
  Toaster,
  ToastTitle,
  useId,
  useToastController,
} from "@fluentui/react-components";
import { confirmStore, noticeStore } from "./store";

export function ConfirmDialog() {
  const request = confirmStore.use();
  const answer = (value: boolean) => {
    confirmStore.set(null);
    request?.resolve(value);
  };
  return (
    <Dialog
      open={request !== null}
      modalType="alert"
      onOpenChange={(_, d) => {
        if (!d.open) answer(false);
      }}
    >
      <DialogSurface>
        <DialogBody>
          <DialogTitle>{request?.title}</DialogTitle>
          <DialogContent>
            <Text block>{request?.body}</Text>
            {request?.code != null && (
              <pre
                data-testid="confirm-code"
                style={{
                  maxHeight: "40vh",
                  overflow: "auto",
                  fontFamily: '"SF Mono", Menlo, monospace',
                  fontSize: 11,
                }}
              >
                {request.code}
              </pre>
            )}
          </DialogContent>
          <DialogActions>
            <Button onClick={() => answer(false)}>{request?.cancel}</Button>
            <Button appearance="primary" onClick={() => answer(true)}>
              {request?.confirm}
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

export function Notices() {
  const toasterId = useId("toaster");
  const { dispatchToast } = useToastController(toasterId);
  const notice = noticeStore.use();
  useEffect(() => {
    if (!notice) return;
    dispatchToast(
      <Toast>
        <ToastTitle>{notice.text}</ToastTitle>
      </Toast>,
      { intent: notice.intent, position: "bottom" },
    );
  }, [notice, dispatchToast]);
  return <Toaster toasterId={toasterId} />;
}
