type Email = `${string}@${string}.${string}`;

type EmailBatchEntry = { email: Email; rowNum: number };

type EmailRow = [Email, `Sent` | `sent` | `${string}` | undefined | null];
