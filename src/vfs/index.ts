import { Bash, type CustomCommand, InMemoryFs } from "just-bash/browser";

let fs: InMemoryFs | null = null;
let bash: Bash | null = null;

let staticFiles: Record<string, string> = {};
let customCommandsFactory: (() => CustomCommand[]) | null = null;

export function setStaticFiles(files: Record<string, string>): void {
  staticFiles = files;
}

export function setCustomCommands(factory: () => CustomCommand[]): void {
  customCommandsFactory = factory;
}

export function getVfs(): InMemoryFs {
  if (!fs) {
    fs = new InMemoryFs({
      "/home/user/uploads/.keep": "",
      ...staticFiles,
    });
  }

  return fs;
}

export function getBash(): Bash {
  if (!bash) {
    bash = new Bash({
      fs: getVfs(),
      cwd: "/home/user",
      customCommands: customCommandsFactory?.() ?? [],
    });
  }

  return bash;
}

export function resetVfs(): void {
  fs = null;
  bash = null;
}

export async function writeFile(path: string, content: string | Uint8Array): Promise<void> {
  const vfs = getVfs();
  const fullPath = path.startsWith("/") ? path : `/home/user/uploads/${path}`;
  const dir = fullPath.substring(0, fullPath.lastIndexOf("/"));

  if (dir && dir !== "/") {
    try {
      await vfs.mkdir(dir, { recursive: true });
    } catch {
    }
  }

  await vfs.writeFile(fullPath, content);
}

export async function readFile(path: string): Promise<string> {
  const vfs = getVfs();
  const fullPath = path.startsWith("/") ? path : `/home/user/uploads/${path}`;
  return vfs.readFile(fullPath);
}

export async function listUploads(): Promise<string[]> {
  const vfs = getVfs();

  try {
    const entries = await vfs.readdir("/home/user/uploads");
    return entries.filter((entry) => entry !== ".keep");
  } catch {
    return [];
  }
}
