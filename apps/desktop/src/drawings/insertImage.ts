import { createDrawingObjectId, type Anchor, type DrawingObject, type ImageEntry, type ImageStore } from "./types";

export async function insertImageFromFile(
  file: File,
  opts: {
    imageId: string;
    anchor: Anchor;
    /**
     * @deprecated Drawing object ids must be globally unique across collaborators; callers should not
     *             pass incrementing counters anymore. This field is ignored.
     */
    nextObjectId?: number;
    objects: DrawingObject[];
    images: ImageStore;
  },
): Promise<{ objects: DrawingObject[]; image: ImageEntry }> {
  const bytes = new Uint8Array(await file.arrayBuffer());
  const mimeType = file.type || guessMimeType(file.name);
  const image: ImageEntry = { id: opts.imageId, bytes, mimeType };
  opts.images.set(image);

  const object: DrawingObject = {
    id: createDrawingObjectId(),
    kind: { type: "image", imageId: image.id },
    anchor: opts.anchor,
    zOrder: opts.objects.length,
  };

  return { objects: [...opts.objects, object], image };
}

function guessMimeType(name: string): string {
  const ext = name.split(".").pop()?.toLowerCase();
  switch (ext) {
    case "png":
      return "image/png";
    case "jpg":
    case "jpeg":
      return "image/jpeg";
    case "gif":
      return "image/gif";
    case "bmp":
      return "image/bmp";
    default:
      return "application/octet-stream";
  }
}
