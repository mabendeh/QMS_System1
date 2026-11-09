import "server-only";
import { createHash, randomBytes } from "node:crypto";
import { mkdir, writeFile, readFile } from "node:fs/promises";
import path from "node:path";

export const MAX_UPLOAD_BYTES = 10 * 1024 * 1024;
const ALLOWED: Record<string, { ext: string; magic: (b: Buffer) => boolean }> = {
  "application/pdf": { ext: "pdf", magic: (b) => b.subarray(0, 5).toString() === "%PDF-" },
  "image/png": { ext: "png", magic: (b) => b.subarray(0, 8).equals(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])) },
  "image/jpeg": { ext: "jpg", magic: (b) => b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff },
  "image/webp": { ext: "webp", magic: (b) => b.subarray(0, 4).toString() === "RIFF" && b.subarray(8, 12).toString() === "WEBP" },
};

export function uploadDir() {
  return path.resolve(process.env.UPLOAD_DIR || "./var/uploads");
}

export class UploadError extends Error {}

/** Validates type by content (not just extension), size, and stores under a random name. */
export async function storeUpload(file: File) {
  if (!file || file.size === 0) throw new UploadError("The file is empty.");
  if (file.size > MAX_UPLOAD_BYTES) throw new UploadError("Files must be 10 MB or smaller.");
  const buf = Buffer.from(await file.arrayBuffer());
  const kind = Object.entries(ALLOWED).find(([, v]) => v.magic(buf));
  if (!kind) throw new UploadError("Only PDF, PNG, JPEG or WEBP files are accepted.");
  const [mimeType, { ext }] = kind;
  const storedName = `${randomBytes(16).toString("hex")}.${ext}`;
  await mkdir(uploadDir(), { recursive: true });
  await writeFile(path.join(uploadDir(), storedName), buf, { mode: 0o600 });
  return {
    storedName,
    mimeType,
    sizeBytes: buf.length,
    sha256: createHash("sha256").update(buf).digest("hex"),
    originalName: (file.name || `evidence.${ext}`).replace(/[^\w.\- ()]/g, "_").slice(0, 120),
  };
}

export async function readUpload(storedName: string) {
  if (!/^[a-f0-9]{32}\.(pdf|png|jpg|webp)$/.test(storedName)) throw new UploadError("Invalid file reference.");
  return readFile(path.join(uploadDir(), storedName));
}
