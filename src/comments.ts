export type CommentStatusFilter = "open" | "resolved" | "all";

export type CommentReply = {
  id: string;
  author: string;
  body: string;
  createdAt: string;
  updatedAt: string;
};

export type CommentData = {
  author: string;
  body: string;
  createdAt: string;
  updatedAt: string;
  resolved: boolean;
  color: string;
  title?: string;
  replies: CommentReply[];
};

export function createCommentData(now = new Date(), author = "You"): CommentData {
  const timestamp = now.toISOString();
  return {
    author,
    body: "",
    createdAt: timestamp,
    updatedAt: timestamp,
    resolved: false,
    color: "#facc15",
    title: "Comment",
    replies: [],
  };
}

export function createCommentReply(body: string, now = new Date(), author = "You"): CommentReply {
  const timestamp = now.toISOString();
  return {
    id: crypto.randomUUID(),
    author,
    body,
    createdAt: timestamp,
    updatedAt: timestamp,
  };
}

export function formatCommentDate(value?: string) {
  if (!value) return "";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "";
  return new Intl.DateTimeFormat(undefined, {
    dateStyle: "medium",
    timeStyle: "short",
  }).format(date);
}

export function getCommentPreview(body: string) {
  const normalized = body.trim().replace(/\s+/g, " ");
  if (!normalized) return "No comment text yet.";
  return normalized.length > 86 ? `${normalized.slice(0, 83)}...` : normalized;
}

export function commentMatchesFilter(resolved: boolean, filter: CommentStatusFilter) {
  if (filter === "all") return true;
  if (filter === "resolved") return resolved;
  return !resolved;
}

export function wrapText(text: string, maxLength: number) {
  const words = text.trim().split(/\s+/).filter(Boolean);
  if (words.length === 0) return [""];

  const lines: string[] = [];
  let line = "";

  for (const word of words) {
    const next = line ? `${line} ${word}` : word;
    if (next.length > maxLength && line) {
      lines.push(line);
      line = word;
    } else {
      line = next;
    }
  }

  if (line) lines.push(line);
  return lines;
}
