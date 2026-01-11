import { stableStringify } from "./cache/key.js";
import { PqDateTimeZone, PqDecimal, PqDuration, PqTime } from "./values.js";

/**
 * @param {Uint8Array} bytes
 * @returns {string}
 */
function bytesToBase64(bytes) {
  if (typeof Buffer !== "undefined") {
    return Buffer.from(bytes).toString("base64");
  }
  let binary = "";
  for (let i = 0; i < bytes.length; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  // eslint-disable-next-line no-undef
  return btoa(binary);
}

/**
 * Compute a deterministic string key for a JS value.
 *
 * This is used anywhere we need "value equality" semantics for non-primitive
 * values (e.g. `Date` objects in join keys) since JS `Map` keys compare objects
 * by identity.
 *
 * @param {unknown} value
 * @returns {string}
 */
export function valueKey(value) {
  if (value == null) return "nil";

  const type = typeof value;
  switch (type) {
    case "string":
      return `str:${value}`;
    case "boolean":
      return value ? "bool:1" : "bool:0";
    case "number": {
      // Match JS `Map`/`Set` semantics: NaN is equal to NaN; -0 is equal to 0.
      if (Number.isNaN(value)) return "num:NaN";
      if (value === Infinity) return "num:Infinity";
      if (value === -Infinity) return "num:-Infinity";
      if (Object.is(value, -0)) return "num:0";
      return `num:${String(value)}`;
    }
    case "bigint":
      return `bigint:${value.toString()}`;
    case "undefined":
      // Keep `null`/`undefined` compatible for join keys.
      return "nil";
    case "symbol":
      return `symbol:${String(value)}`;
    case "function":
      return `function:${value.name || "<anonymous>"}`;
    case "object": {
      if (value instanceof Date) {
        const time = value.getTime();
        return Number.isNaN(time) ? "date:NaN" : `date:${time}`;
      }

      if (value instanceof PqDecimal) {
        return `decimal:${value.value}`;
      }

      if (value instanceof PqTime) {
        return `time:${value.toString()}`;
      }

      if (value instanceof PqDuration) {
        return `duration:${value.toString()}`;
      }

      if (value instanceof PqDateTimeZone) {
        return `datetimezone:${value.toString()}`;
      }

      if (value instanceof Uint8Array) {
        return `binary:${bytesToBase64(value)}`;
      }

      if (Array.isArray(value)) {
        return `array:${stableStringify(value)}`;
      }

      return `object:${stableStringify(value)}`;
    }
    default: {
      /** @type {never} */
      const exhausted = type;
      return `unknown:${String(exhausted)}`;
    }
  }
}
