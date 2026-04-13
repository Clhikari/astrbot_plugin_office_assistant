import * as fs from "node:fs";
import * as path from "node:path";

function contractSearchRoots(): string[] {
  return [
    path.resolve(__dirname, "../../../shared_contracts"),
    path.resolve(process.cwd(), "shared_contracts"),
    path.resolve(process.cwd(), "../shared_contracts"),
  ];
}

function resolveSharedContractPath(name: string): string {
  for (const root of contractSearchRoots()) {
    const candidate = path.join(root, name);
    if (fs.existsSync(candidate)) {
      return candidate;
    }
  }
  throw new Error(`Shared contract not found: ${name}`);
}

export function readSharedContract<T>(name: string): T {
  return JSON.parse(
    fs.readFileSync(resolveSharedContractPath(name), "utf8"),
  ) as T;
}
