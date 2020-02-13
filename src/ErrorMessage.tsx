import React from "react";
import { Alert } from "reactstrap";

export interface IErrorMessageProps {
  debug: string;
  message: string;
}

const ErrorMessage = (props: IErrorMessageProps) => {
  let debug = null;
  if (props.debug) {
    debug = (
      <pre className="alert-pre border bg-light p-2">
        <code>{props.debug}</code>
      </pre>
    );
  }
  return (
    <Alert color="danger">
      <p className="mb-3">{props.message}</p>
      {debug}
    </Alert>
  );
};
export default ErrorMessage;
