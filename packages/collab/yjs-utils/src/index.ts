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

export type DocTypeConstructors = { Map: new () => any; Array: new () => any; Text: new () => any };

/**
 * Return constructors for Y.Map/Y.Array/Y.Text that match the module instance used to
 * create `doc`.
 *
 * In pnpm workspaces it is possible to load both the ESM + CJS builds of Yjs in the
 * same process (for example via y-websocket). Yjs types cannot be moved across module
 * instances; using constructors derived from the target doc avoids "Unexpected content
 * type" errors when cloning values into that doc.
 */
export function getDocTypeConstructors(doc: unknown): DocTypeConstructors {
  const DocCtor = (doc as any)?.constructor as (new () => any) | undefined;
  if (typeof DocCtor !== "function") {
    return { Map: Y.Map, Array: Y.Array, Text: Y.Text };
  }

  try {
    const probe = new DocCtor();
    const ctors: DocTypeConstructors = {
      Map: probe.getMap("__ctor_probe_map").constructor,
      Array: probe.getArray("__ctor_probe_array").constructor,
      Text: probe.getText("__ctor_probe_text").constructor,
    };
    probe.destroy?.();
    return ctors;
  } catch {
    return { Map: Y.Map, Array: Y.Array, Text: Y.Text };
  }
}

const patchedItemConstructors = new WeakSet<Function>();
const patchedAbstractTypeConstructors = new WeakSet<Function>();
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

/**
 * Patch the prototype chain for a foreign Yjs type so it passes `instanceof Y.AbstractType`
 * checks against this module instance.
 *
 * In mixed-module environments (ESM + CJS), a document can contain types created by a
 * different `yjs` module instance. Yjs' UndoManager performs `instanceof AbstractType`
 * checks when working with scopes; if these checks fail, it can emit warnings like
 * `[yjs#509] Not same Y.Doc` and behave incorrectly.
 *
 * This helper is best-effort and intentionally avoids breaking `instanceof` checks in
 * the foreign module instance by inserting the local `AbstractType.prototype` into the
 * existing prototype chain rather than replacing it outright.
 */
export function patchForeignAbstractTypeConstructor(type: unknown): void {
  if (!type || typeof type !== "object") return;
  if (!isYAbstractType(type)) return;
  if (type instanceof Y.AbstractType) return;

  const ctor = (type as any).constructor as Function | undefined;
  if (!ctor || ctor === Y.AbstractType) return;
  if (patchedAbstractTypeConstructors.has(ctor)) return;
  patchedAbstractTypeConstructors.add(ctor);

  try {
    const baseProto = Object.getPrototypeOf((ctor as any).prototype);
    // `ctor.prototype` is usually a concrete type prototype (e.g. YMap.prototype),
    // whose base prototype is the foreign AbstractType prototype. Patch that base
    // prototype so the local AbstractType prototype is also in the chain.
    if (baseProto && baseProto !== Object.prototype) {
      Object.setPrototypeOf(baseProto, Y.AbstractType.prototype);
    } else {
      Object.setPrototypeOf((ctor as any).prototype, Y.AbstractType.prototype);
    }
  } catch {
    // Best-effort: if we can't patch (frozen prototypes, etc), behave like upstream Yjs.
  }
}

export function patchForeignItemConstructor(item: unknown): void {
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

/**
 * Deep-clone a Yjs value from one document into freshly-constructed Yjs types.
 *
 * This is used when copying state between docs (e.g. snapshot/restore) since Yjs
 * types cannot be moved between documents or between different `yjs` module
 * instances (ESM vs CJS).
 *
 * When `constructors` is provided, it is used to create new Yjs types compatible
 * with the target doc's module instance (see `getDocTypeConstructors`).
 */
export function cloneYjsValue(value: any, constructors: { Map?: new () => any; Array?: new () => any; Text?: new () => any } | null = null): any {
  const map = getYMap(value);
  if (map) {
    const MapCtor = constructors?.Map ?? (map as any).constructor;
    const out = new (MapCtor as any)();
    const prelim = (map as any)?.doc == null ? (map as any)?._prelimContent : null;
    const keys =
      prelim instanceof Map ? Array.from(prelim.keys()).sort() : Array.from(map.keys()).sort();
    for (const key of keys) {
      const nextValue = prelim instanceof Map ? prelim.get(key) : map.get(key);
      out.set(key, cloneYjsValue(nextValue, constructors));
    }

    // Yjs intentionally warns on reading unintegrated types (doc=null) via `get()`. For cloning
    // and test usage we want the returned value to behave like a normal map before insertion,
    // so fall back to `_prelimContent` while the type is unintegrated.
    const outPrelim = (out as any)?._prelimContent;
    if ((out as any)?.doc == null && outPrelim instanceof Map && typeof (out as any).get === "function") {
      const originalGet = (out as any).get as (key: any) => any;
      try {
        (out as any).get = function patchedGet(key: any) {
          if ((this as any)?.doc == null && (this as any)?._prelimContent instanceof Map) {
            return (this as any)._prelimContent.get(key);
          }
          return originalGet.call(this, key);
        };
      } catch {
        // Best-effort.
      }
    }
    return out;
  }

  const array = getYArray(value);
  if (array) {
    const ArrayCtor = constructors?.Array ?? (array as any).constructor;
    const out = new (ArrayCtor as any)();
    const prelim = (array as any)?.doc == null && Array.isArray((array as any)?._prelimContent) ? (array as any)._prelimContent : null;
    for (const item of prelim ?? array.toArray()) {
      out.push([cloneYjsValue(item, constructors)]);
    }
    return out;
  }

  const text = getYText(value);
  if (text) {
    const TextCtor = constructors?.Text ?? (text as any).constructor;
    const out = new (TextCtor as any)();

    const structuredCloneFn = (globalThis as any).structuredClone as ((input: unknown) => unknown) | undefined;
    const safeClone = <T>(input: T): T => {
      if (typeof structuredCloneFn !== "function") return input;
      try {
        return structuredCloneFn(input) as T;
      } catch {
        return input;
      }
    };

    const cloneDelta = (delta: any[]): any[] => {
      return delta.map((op) => {
        const next: any = { ...op };
        if ("insert" in next) {
          const insert = next.insert;
          if (insert && typeof insert === "object") {
            // Y.Text embeds can include nested Yjs types; clone them into the target constructors.
            const yEmbed = getYMap(insert) || getYArray(insert) || getYText(insert);
            if (yEmbed) {
              next.insert = cloneYjsValue(insert, constructors);
            } else if (typeof structuredCloneFn === "function") {
              try {
                next.insert = structuredCloneFn(insert);
              } catch {
                // Fall back to returning the embed as-is.
              }
            }
          }
        }
        if (next.attributes && typeof next.attributes === "object" && typeof structuredCloneFn === "function") {
          try {
            next.attributes = structuredCloneFn(next.attributes);
          } catch {
            // ignore
          }
        }
        return next;
      });
    };

    // Preserve formatting by cloning the Y.Text delta instead of just its plain string.
    //
    // Note: unintegrated Y.Text instances store edits in `_pending`; calling `toDelta()` returns
    // an empty delta until the text is integrated into a Doc. To avoid mutating the source
    // (integrating it into a temporary doc), replay the queued operations by intercepting
    // method calls, then apply them to the clone.
    if ((text as any)?.doc != null) {
      out.applyDelta(cloneDelta(text.toDelta()));
      return out;
    }

    const pending = (text as any)?._pending;
    if (Array.isArray(pending)) {
      const ops: Array<{ kind: string; args: any[] }> = [];
      const record = (kind: string, args: any[]): void => {
        ops.push({ kind, args });
      };

      const original = {
        insert: (text as any).insert,
        delete: (text as any).delete,
        format: (text as any).format,
        applyDelta: (text as any).applyDelta,
        insertEmbed: (text as any).insertEmbed,
        removeAttribute: (text as any).removeAttribute,
        setAttribute: (text as any).setAttribute,
      };

      try {
        (text as any).insert = (index: number, str: string, attributes?: any) => record("insert", [index, str, attributes]);
        (text as any).delete = (index: number, length: number) => record("delete", [index, length]);
        (text as any).format = (index: number, length: number, attributes?: any) => record("format", [index, length, attributes]);
        (text as any).applyDelta = (delta: any) => record("applyDelta", [delta]);
        (text as any).insertEmbed = (index: number, embed: any, attributes?: any) => record("insertEmbed", [index, embed, attributes]);
        (text as any).removeAttribute = (attributeName: string) => record("removeAttribute", [attributeName]);
        (text as any).setAttribute = (attributeName: string, attributeValue: any) =>
          record("setAttribute", [attributeName, attributeValue]);

        for (const fn of pending) {
          if (typeof fn !== "function") continue;
          try {
            fn();
          } catch {
            // Best-effort.
          }
        }
      } finally {
        (text as any).insert = original.insert;
        (text as any).delete = original.delete;
        (text as any).format = original.format;
        (text as any).applyDelta = original.applyDelta;
        (text as any).insertEmbed = original.insertEmbed;
        (text as any).removeAttribute = original.removeAttribute;
        (text as any).setAttribute = original.setAttribute;
      }

      for (const op of ops) {
        switch (op.kind) {
          case "insert": {
            const [index, str, attributes] = op.args;
            (out as any).insert(index, str, attributes ? safeClone(attributes) : attributes);
            break;
          }
          case "delete": {
            const [index, length] = op.args;
            (out as any).delete(index, length);
            break;
          }
          case "format": {
            const [index, length, attributes] = op.args;
            (out as any).format(index, length, attributes ? safeClone(attributes) : attributes);
            break;
          }
          case "applyDelta": {
            const [delta] = op.args;
            (out as any).applyDelta(cloneDelta(Array.isArray(delta) ? delta : []));
            break;
          }
          case "insertEmbed": {
            const [index, embed, attributes] = op.args;
            (out as any).insertEmbed(index, cloneYjsValue(embed, constructors), attributes ? safeClone(attributes) : attributes);
            break;
          }
          case "removeAttribute": {
            const [attributeName] = op.args;
            (out as any).removeAttribute(attributeName);
            break;
          }
          case "setAttribute": {
            const [attributeName, attributeValue] = op.args;
            (out as any).setAttribute(attributeName, cloneYjsValue(attributeValue, constructors));
            break;
          }
          default:
            break;
        }
      }
    } else if (typeof (text as any)._integrate === "function") {
      // Best-effort fallback: some preliminary Y.Text values may not expose `_pending` in a usable
      // way (e.g. across module instances or after prototype patching). In that case, try to
      // materialize the pending operations by integrating into a temporary doc and cloning the
      // resulting delta.
      let delta: any[] = [];
      try {
        const doc = new Y.Doc();
        try {
          (text as any)._integrate(doc, null);
          delta = text.toDelta();
        } finally {
          doc.destroy?.();
        }
      } catch {
        delta = [];
      }
      out.applyDelta(cloneDelta(delta));
    }
    return out;
  }

  if (Array.isArray(value)) {
    return value.map((item) => cloneYjsValue(item, constructors));
  }

  if (value && typeof value === "object") {
    const structuredCloneFn = (globalThis as any).structuredClone as ((input: unknown) => unknown) | undefined;
    if (typeof structuredCloneFn === "function") {
      try {
        return structuredCloneFn(value);
      } catch {
        // Ignore: fall through.
      }
    }
  }

  return value;
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
      // `yArr.toArray()` allocates a fresh JS array. Mutate it in-place to avoid
      // the additional array allocation that `.map()` would create.
      const out = yArr.toArray();
      for (let i = 0; i < out.length; i += 1) {
        out[i] = yjsValueToJson(out[i]);
      }
      return out;
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
