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
        n.parent = t;
      }
    });
  }

  for (let n = existing?._start ?? null; n !== null; n = n.right) {
    n.parent = t;
  }

  doc.share.set(name, t);
  if (typeof t._integrate === "function") {
    t._integrate(doc as any, null);
  }
  return t as T;
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
