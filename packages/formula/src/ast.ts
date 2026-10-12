import type { ErrorCode } from "./values";
import type { RefToken } from "./lexer";

export type BinaryOp = "+" | "-" | "*" | "/" | "^" | "&" | "=" | "<>" | "<" | ">" | "<=" | ">=" | ":" | " ";

export type Node =
  | { type: "num"; value: number }
  | { type: "str"; value: string }
  | { type: "bool"; value: boolean }
  | { type: "err"; code: ErrorCode }
  | { type: "ref"; ref: RefToken }
  | { type: "name"; name: string }
  | { type: "func"; name: string; args: Node[] }
  | { type: "unary"; op: "-" | "+"; arg: Node }
  | { type: "postfix"; op: "%"; arg: Node }
  | { type: "binary"; op: BinaryOp; left: Node; right: Node }
  | { type: "union"; items: Node[] }
  | { type: "array"; rows: Node[][] }
  | { type: "missing" };

export class ParseError extends Error {
  constructor(
    message: string,
    readonly pos: number,
  ) {
    super(message);
  }
}

/** AST'deki tüm başvuru ve ad düğümlerini gezer. */
export function walk(node: Node, visit: (n: Node) => void): void {
  visit(node);
  switch (node.type) {
    case "func":
      node.args.forEach((a) => walk(a, visit));
      break;
    case "unary":
    case "postfix":
      walk(node.arg, visit);
      break;
    case "binary":
      walk(node.left, visit);
      walk(node.right, visit);
      break;
    case "union":
      node.items.forEach((a) => walk(a, visit));
      break;
    case "array":
      node.rows.forEach((r) => r.forEach((a) => walk(a, visit)));
      break;
  }
}
