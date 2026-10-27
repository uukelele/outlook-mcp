export class OutlookError extends Error {
  constructor(message: string, readonly code?: string) {
    super(message);
    this.name = new.target.name;
  }
}

/** Credentials are missing, invalid, or expired. */
export class AuthError extends OutlookError {}

/** The service returned a non-success HTTP status or a non-JSON body. */
export class OwaHttpError extends OutlookError {
  constructor(
    message: string,
    readonly status: number,
    readonly action: string,
    code?: string,
  ) {
    super(message, code);
  }
}

/** HTTP 200, but the response message reports ResponseClass "Error". */
export class OwaResponseError extends OutlookError {
  constructor(
    readonly action: string,
    readonly responseCode: string,
    readonly messageText: string | undefined,
  ) {
    super(describeResponseCode(action, responseCode, messageText), responseCode);
  }
}

const HINTS: Record<string, string> = {
  ErrorItemNotFound: "the item was not found; it may have been moved or deleted, or the ID is wrong",
  ErrorFolderNotFound: "the folder was not found; use list_folders to find valid folder IDs",
  ErrorInvalidIdMalformed: "the ID is malformed",
  ErrorAccessDenied: "access denied",
  ErrorInvalidRecipients: "one or more recipients are invalid",
  ErrorChangeKeyRequiredForWriteOperations: "the item changed on the server; fetch it again and retry",
  ErrorIrresolvableConflict: "the item was modified concurrently; fetch it again and retry",
  ErrorQuotaExceeded: "the mailbox is over its storage quota",
  ErrorMessageSizeExceeded: "the message is larger than the mailbox allows",
  ErrorFolderExists: "a folder with that name already exists here",
  ErrorCannotDeleteObject: "the server refused to delete this object",
};

export function describeResponseCode(action: string, code: string, messageText?: string): string {
  const hint = HINTS[code];
  const detail = [hint, messageText].filter(Boolean).join(" — ");
  return `${action} failed (${code})${detail ? `: ${detail}` : ""}`;
}

export function errorMessage(error: unknown): string {
  if (error instanceof Error) return error.message;
  return String(error);
}
