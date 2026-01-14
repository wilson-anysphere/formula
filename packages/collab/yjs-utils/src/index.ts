import * as Y from "yjs";

/**
 * Duck-type helpers to tolerate multiple `yjs` module instances (ESM/CJS) and
 * constructor renaming by bundlers.
 *
 * These utilities intentionally avoid relying on `instanceof` alone, because a
 * single document may contain Yjs types created by a different module instance
 * (e.g. y-websocket using CJS `require("yjs")` while the app uses ESM imports).
 */

export function getYMap(value: unknown): any | null {
  if (value instanceof Y.Map) return value;
  if (!value || typeof value !== "object") return null;
  const maybe = value as any;
  if (typeof maybe.get !== "function") return null;
  if (typeof maybe.set !== "function") return null;
  if (typeof maybe.delete !== "function") return null;
  if (typeof maybe.keys !== "function") return null;
  if (typeof maybe.forEach !== "function") return null;
  // Plain JS Maps also have get/set/delete/keys/forEach; require Yjs observer APIs.
  if (typeof maybe.observeDeep !== "function") return null;
  if (typeof maybe.unobserveDeep !== "function") return null;
  return maybe;
}

export function getYArray(value: unknown): any | null {
  if (value instanceof Y.Array) return value;
  if (!value || typeof value !== "object") return null;
  const maybe = value as any;
  // See `getYMap` above for rationale.
  if (typeof maybe.get !== "function") return null;
  if (typeof maybe.toArray !== "function") return null;
  if (typeof maybe.push !== "function") return null;
  if (typeof maybe.delete !== "function") return null;
  if (typeof maybe.observeDeep !== "function") return null;
  if (typeof maybe.unobserveDeep !== "function") return null;
  return maybe;
}

export function getYText(value: unknown): any | null {
  if (value instanceof Y.Text) return value;
  if (!value || typeof value !== "object") return null;
  const maybe = value as any;
  // See `getYMap` above for rationale.
  if (typeof maybe.toString !== "function") return null;
  if (typeof maybe.toDelta !== "function") return null;
  if (typeof maybe.applyDelta !== "function") return null;
  if (typeof maybe.insert !== "function") return null;
  if (typeof maybe.delete !== "function") return null;
  if (typeof maybe.observeDeep !== "function") return null;
  if (typeof maybe.unobserveDeep !== "function") return null;
  return maybe;
}

export function isYAbstractType(value: unknown): boolean {
  if (value instanceof Y.AbstractType) return true;
  if (!value || typeof value !== "object") return false;
  const maybe = value as any;
  if (typeof maybe.observeDeep !== "function") return false;
  if (typeof maybe.unobserveDeep !== "function") return false;
  return Boolean(maybe._map instanceof Map || maybe._start || maybe._item || maybe._length != null);
}

const patchedItemConstructors = new WeakSet<Function>();
const patchedContentConstructors = new WeakMap<Function, Function>();

function isYjsItemStruct(value: unknown): value is any {
  if (!value || typeof value !== "object") return false;
  const maybe = value as any;
  if (!("id" in maybe)) return false;
  if (typeof maybe.length !== "number") return false;
  if (!("content" in maybe)) return false;
  if (!("parent" in maybe)) return false;
  if (!("parentSub" in maybe)) return false;
  if (typeof maybe.content?.getContent !== "function") return false;
  return true;
}

function patchForeignItemConstructor(item: unknown): void {
  if (!isYjsItemStruct(item)) return;
  if (item instanceof Y.Item) return;
  const ctor = (item as any).constructor as Function | undefined;
  if (!ctor || ctor === Y.Item) return;
  if (patchedItemConstructors.has(ctor)) return;
  patchedItemConstructors.add(ctor);
  try {
    Object.setPrototypeOf((ctor as any).prototype, Y.Item.prototype);
    (ctor as any).prototype.constructor = Y.Item;
  } catch {
    // Best-effort.
  }
}

function patchForeignContentConstructor(content: unknown): void {
  if (!content || typeof content !== "object") return;
  const ctor = (content as any).constructor as Function | undefined;
  if (!ctor) return;
  if (patchedContentConstructors.has(ctor)) return;

  // Prefer duck typing over constructor names so bundlers that rename constructors
  // still work. Patch only the content types we know are used by Y.Text methods.
  let localCtor: Function | null = null;
  const maybe = content as any;
  if (typeof maybe.str === "string") {
    localCtor = Y.ContentString;
  } else if (typeof maybe.key === "string" && ("value" in maybe || "val" in maybe)) {
    localCtor = Y.ContentFormat;
  } else if ("embed" in maybe) {
    localCtor = Y.ContentEmbed;
  }

  if (!localCtor || ctor === localCtor) {
    patchedContentConstructors.set(ctor, ctor);
    return;
  }

  patchedContentConstructors.set(ctor, localCtor);
  try {
    Object.setPrototypeOf((ctor as any).prototype, (localCtor as any).prototype);
    (ctor as any).prototype.constructor = localCtor;
  } catch {
    // Best-effort.
  }
}

export function replaceForeignRootType<T>(params: { doc: Y.Doc; name: string; existing: any; create: () => T }): T {
  const { doc, name, existing, create } = params;
  // If the whole doc was created by a different Yjs module instance (ESM vs CJS),
  // we cannot safely insert local types into it. Callers should generally avoid
  // invoking this helper in that situation.
  if (!(doc instanceof Y.Doc)) return existing as T;
  const t: any = create();

  // Mirror Yjs' own Doc.get conversion logic for AbstractType placeholders, but
  // also support roots instantiated by a different Yjs module instance (e.g.
  // CJS `require("yjs")` or CJS `applyUpdate`).
  t._map = existing?._map;
  t._start = existing?._start;
  t._length = existing?._length;

  // Update parent pointers so future updates can resolve the correct root key via
  // `findRootTypeKey` when encoding.
  const map = existing?._map;
  if (map instanceof Map) {
    map.forEach((item: any) => {
      for (let n = item; n !== null; n = n.left) {
        patchForeignItemConstructor(n);
        patchForeignContentConstructor(n?.content);
        n.parent = t;
      }
    });
  }

  for (let n = existing?._start ?? null; n !== null; n = n.right) {
    patchForeignItemConstructor(n);
    patchForeignContentConstructor(n?.content);
    n.parent = t;
  }

  doc.share.set(name, t);
  if (typeof t._integrate === "function") {
    t._integrate(doc as any, null);
  }

  // Y.Text special case: Yjs' rich-text helpers use constructor equality checks
  // (`n.content.constructor === ContentString`) to detect content types.
  //
  // When the existing root was created by a *different* Yjs module instance,
  // each `Item.content` object is also from that module instance. If we only
  // replace the root type, the new local Y.Text will see "foreign" Content*
  // instances and treat them as unknown, causing `toString()` / `toDelta()` to
  // return empty results.
  //
  // Fix: patch the prototype of known Content* objects based on their `getRef()`
  // so `content.constructor` matches the local Content* constructors.
  if (t instanceof Y.Text) {
    const protosByRef = getYTextContentPrototypesByRef();
    for (let n = t._start ?? null; n !== null; n = n.right) {
      const content = n?.content;
      if (!content || typeof content !== "object" || typeof content.getRef !== "function") continue;
      const ref = content.getRef();
      if (typeof ref !== "number") continue;
      const proto = protosByRef.get(ref);
      if (!proto) continue;
      if (Object.getPrototypeOf(content) === proto) continue;
      try {
        Object.setPrototypeOf(content, proto);
      } catch {
        // Ignore: non-extensible content objects (unexpected).
      }
    }
  }
  return t as T;
}

let yTextContentPrototypesByRef: Map<number, object> | null = null;

function getYTextContentPrototypesByRef(): Map<number, object> {
  if (yTextContentPrototypesByRef) return yTextContentPrototypesByRef;

  /**
   * Extract `Item.content` prototypes from a probe Y.Text by walking its internal item list.
   * We key by Yjs' internal content "ref" ids (`content.getRef()`).
   *
   * @param {Y.Text} text
   * @param {Map<number, object>} out
   */
  const collectProtos = (text: Y.Text, out: Map<number, object>) => {
    for (let n = (text as any)?._start ?? null; n !== null; n = n.right) {
      const content = n?.content;
      if (!content || typeof content !== "object" || typeof content.getRef !== "function") continue;
      const ref = content.getRef();
      if (typeof ref !== "number" || out.has(ref)) continue;
      out.set(ref, Object.getPrototypeOf(content));
    }
  };

  const doc = new Y.Doc();
  const protos = new Map<number, object>();

  // Populate the common Y.Text content types:
  // - 4: ContentString (plain text)
  // - 5: ContentEmbed (inserted embed objects)
  // - 6: ContentFormat (formatting attributes)
  // - 7: ContentType (embedded Yjs types)
  const tString = doc.getText("__content_proto_string");
  tString.insert(0, "x");
  collectProtos(tString, protos);

  const tEmbed = doc.getText("__content_proto_embed");
  tEmbed.insertEmbed(0, { foo: "bar" });
  collectProtos(tEmbed, protos);

  const tType = doc.getText("__content_proto_type");
  tType.insertEmbed(0, new Y.Map());
  collectProtos(tType, protos);

  const tFormat = doc.getText("__content_proto_format");
  tFormat.insert(0, "x");
  tFormat.format(0, 1, { bold: true });
  collectProtos(tFormat, protos);

  doc.destroy();

  yTextContentPrototypesByRef = protos;
  return protos;
}

export function getMapRoot<T = unknown>(doc: Y.Doc, name: string): Y.Map<T> {
  const existing = doc.share.get(name);
  if (!existing) return doc.getMap<T>(name);

  const map = getYMap(existing);
  if (map) {
    return map instanceof Y.Map ? (map as Y.Map<T>) : (replaceForeignRootType({ doc, name, existing: map, create: () => new Y.Map() }) as any);
  }

  const array = getYArray(existing);
  if (array) {
    throw new Error(`Yjs root schema mismatch for "${name}": expected a Y.Map but found a Y.Array`);
  }

  const text = getYText(existing);
  if (text) {
    throw new Error(`Yjs root schema mismatch for "${name}": expected a Y.Map but found a Y.Text`);
  }

  // `instanceof Y.AbstractType` is not sufficient to detect whether the placeholder
  // root was created by *this* Yjs module instance. In mixed-module environments
  // (ESM + CJS), other parts of the system (e.g. collaborative undo) patch foreign
  // prototype chains so foreign types pass `instanceof` checks.
  //
  // When that happens, calling `doc.getMap(name)` would throw "different constructor"
  // for foreign placeholders. Use constructor identity instead.
  if (existing instanceof Y.AbstractType && (existing as any).constructor === Y.AbstractType) {
    return doc.getMap<T>(name);
  }

  if (isYAbstractType(existing)) {
    if (doc instanceof Y.Doc) {
      return replaceForeignRootType({ doc, name, existing, create: () => new Y.Map() }) as any;
    }
    return doc.getMap<T>(name);
  }

  throw new Error(`Unsupported Yjs root type for "${name}": ${existing?.constructor?.name ?? typeof existing}`);
}

export function getArrayRoot<T = unknown>(doc: Y.Doc, name: string): Y.Array<T> {
  const existing = doc.share.get(name);
  if (!existing) return doc.getArray<T>(name);

  const arr = getYArray(existing);
  if (arr) {
    return arr instanceof Y.Array ? (arr as Y.Array<T>) : (replaceForeignRootType({ doc, name, existing: arr, create: () => new Y.Array() }) as any);
  }

  const map = getYMap(existing);
  if (map) {
    throw new Error(`Yjs root schema mismatch for "${name}": expected a Y.Array but found a Y.Map`);
  }

  const text = getYText(existing);
  if (text) {
    throw new Error(`Yjs root schema mismatch for "${name}": expected a Y.Array but found a Y.Text`);
  }

  if (existing instanceof Y.AbstractType && (existing as any).constructor === Y.AbstractType) {
    return doc.getArray<T>(name);
  }

  if (isYAbstractType(existing)) {
    if (doc instanceof Y.Doc) {
      return replaceForeignRootType({ doc, name, existing, create: () => new Y.Array() }) as any;
    }
    return doc.getArray<T>(name);
  }

  throw new Error(`Unsupported Yjs root type for "${name}": ${existing?.constructor?.name ?? typeof existing}`);
}

export function getTextRoot(doc: Y.Doc, name: string): Y.Text {
  const existing = doc.share.get(name);
  if (!existing) return doc.getText(name);

  const text = getYText(existing);
  if (text) {
    return text instanceof Y.Text ? (text as Y.Text) : (replaceForeignRootType({ doc, name, existing: text, create: () => new Y.Text() }) as any);
  }

  const map = getYMap(existing);
  if (map) {
    throw new Error(`Yjs root schema mismatch for "${name}": expected a Y.Text but found a Y.Map`);
  }

  const array = getYArray(existing);
  if (array) {
    throw new Error(`Yjs root schema mismatch for "${name}": expected a Y.Text but found a Y.Array`);
  }

  if (existing instanceof Y.AbstractType && (existing as any).constructor === Y.AbstractType) {
    return doc.getText(name);
  }

  if (isYAbstractType(existing)) {
    if (doc instanceof Y.Doc) {
      return replaceForeignRootType({ doc, name, existing, create: () => new Y.Text() }) as any;
    }
    return doc.getText(name);
  }

  throw new Error(`Unsupported Yjs root type for "${name}": ${existing?.constructor?.name ?? typeof existing}`);
}

function isPlainObject(value: unknown): value is Record<string, any> {
  if (!value || typeof value !== "object") return false;
  if (Array.isArray(value)) return false;
  const proto = Object.getPrototypeOf(value);
  return proto === Object.prototype || proto === null;
}

/**
 * Convert (potentially nested) Yjs values into plain JS objects.
 *
 * Intended for stable comparisons, debug output, and key computation. This is
 * a best-effort clone: unknown/non-serializable objects are returned as-is.
 */
export function yjsValueToJson(value: any): any {
  const text = getYText(value);
  if (text) return text.toString();

  if (value && typeof value === "object") {
    const yArr = getYArray(value);
    if (yArr) {
      return yArr.toArray().map((item: any) => yjsValueToJson(item));
    }

    const yMap = getYMap(value);
    if (yMap) {
      const out: Record<string, any> = {};
      const keys = Array.from(yMap.keys()).sort();
      for (const key of keys) out[String(key)] = yjsValueToJson(yMap.get(key));
      return out;
    }

    if (Array.isArray(value)) return value.map((item) => yjsValueToJson(item));

    if (isPlainObject(value)) {
      const out: Record<string, any> = {};
      const keys = Object.keys(value).sort();
      for (const key of keys) out[key] = yjsValueToJson((value as any)[key]);
      return out;
    }

    // Preserve other objects as-is (clone to avoid accidental mutations).
    if (!Array.isArray(value)) {
      const structuredCloneFn = (globalThis as any).structuredClone as ((input: unknown) => unknown) | undefined;
      if (typeof structuredCloneFn === "function") {
        try {
          return structuredCloneFn(value);
        } catch {
          // Ignore: fall through.
        }
      }
    }
  }

  return value;
}

export const cloneYjsValueToJson = yjsValueToJson;
