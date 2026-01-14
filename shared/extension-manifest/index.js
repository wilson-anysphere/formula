const GLOBAL_KEY = "__formula_extension_manifest__";

const SEMVER_RANGE_KEY = "__formula_semver_range__";

function getSemverRange() {
  // Node / CommonJS: use require when available.
  if (typeof require === "function") {
    try {
      return require("../semver-range");
    } catch {
      // ignore
    }
  }

  // Browser / ESM: semver-range is expected to have registered itself on globalThis.
  try {
    if (typeof globalThis === "undefined") return null;
    return globalThis[SEMVER_RANGE_KEY] || null;
  } catch {
    return null;
  }
}

const semverRange = getSemverRange();
if (!semverRange) {
  throw new Error("shared/extension-manifest: failed to initialize semver-range");
}

const { isValidSemver, satisfies } = semverRange;

function pathExtname(p) {
  const raw = String(p ?? "");
  const lastSlash = Math.max(raw.lastIndexOf("/"), raw.lastIndexOf("\\"));
  const lastDot = raw.lastIndexOf(".");
  if (lastDot === -1 || lastDot < lastSlash) return "";
  return raw.slice(lastDot).toLowerCase();
}

function assertEntrypointExtension(filePath, label, allowedExts) {
  const ext = pathExtname(String(filePath).trim());
  if (!allowedExts.has(ext)) {
    const expected = [...allowedExts].sort().join(", ");
    throw new ManifestError(`${label} entrypoint must end with one of: ${expected} (got ${filePath})`);
  }
}

const VALID_PERMISSIONS = new Set([
  "cells.read",
  "cells.write",
  "sheets.manage",
  "workbook.manage",
  "network",
  "clipboard",
  "storage",
  "ui.panels",
  "ui.commands",
  "ui.menus",
]);

const MAIN_ENTRYPOINT_EXTS = new Set([".cjs", ".js"]);
const ESM_ENTRYPOINT_EXTS = new Set([".js", ".mjs"]);

class ManifestError extends Error {
  constructor(message) {
    super(message);
    this.name = "ManifestError";
  }
}

function assertObject(value, label) {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new ManifestError(`${label} must be an object`);
  }
  return value;
}

function assertString(value, label) {
  if (typeof value !== "string") {
    throw new ManifestError(`${label} must be a non-empty string`);
  }
  const trimmed = value.trim();
  if (trimmed.length === 0) {
    throw new ManifestError(`${label} must be a non-empty string`);
  }
  return trimmed;
}

function assertOptionalString(value, label) {
  if (value === undefined) return undefined;
  if (value === null) return undefined;
  if (typeof value !== "string") throw new ManifestError(`${label} must be a string`);
  return value.trim();
}

function assertArray(value, label) {
  if (value === undefined) return [];
  if (!Array.isArray(value)) {
    throw new ManifestError(`${label} must be an array`);
  }
  return value;
}

function validateCommands(commands) {
  const list = assertArray(commands, "contributes.commands");
  const seen = new Set();
  const normalized = [];
  for (const [idx, cmd] of list.entries()) {
    assertObject(cmd, `contributes.commands[${idx}]`);
    const id = assertString(cmd.command, `contributes.commands[${idx}].command`);
    if (seen.has(id)) throw new ManifestError(`Duplicate command id: ${id}`);
    seen.add(id);
    const title = assertString(cmd.title, `contributes.commands[${idx}].title`);
    const category = assertOptionalString(cmd.category, `contributes.commands[${idx}].category`);
    const icon = assertOptionalString(cmd.icon, `contributes.commands[${idx}].icon`);
    const description = assertOptionalString(cmd.description, `contributes.commands[${idx}].description`);

    const keywordsRaw = cmd.keywords;
    const keywords = keywordsRaw === undefined ? undefined : assertArray(keywordsRaw, `contributes.commands[${idx}].keywords`);
    const normalizedKeywords =
      keywords === undefined
        ? undefined
        : keywords.map((keyword, kIdx) =>
            assertString(keyword, `contributes.commands[${idx}].keywords[${kIdx}]`),
          );

    const normalizedCommand = { ...cmd, command: id, title };
    if (category !== undefined) normalizedCommand.category = category;
    if (icon !== undefined) normalizedCommand.icon = icon;
    if (description !== undefined) normalizedCommand.description = description;
    if (normalizedKeywords !== undefined) normalizedCommand.keywords = normalizedKeywords;
    normalized.push(normalizedCommand);
  }
  return normalized;
}

function validatePanels(panels) {
  const list = assertArray(panels, "contributes.panels");
  const seen = new Set();
  const normalized = [];
  for (const [idx, panel] of list.entries()) {
    assertObject(panel, `contributes.panels[${idx}]`);
    const id = assertString(panel.id, `contributes.panels[${idx}].id`);
    if (seen.has(id)) throw new ManifestError(`Duplicate panel id: ${id}`);
    seen.add(id);
    const title = assertString(panel.title, `contributes.panels[${idx}].title`);
    const icon = assertOptionalString(panel.icon, `contributes.panels[${idx}].icon`);
    const normalizedPanel = { ...panel, id, title };
    if (icon !== undefined) normalizedPanel.icon = icon;
    normalized.push(normalizedPanel);
  }
  return normalized;
}

function validateKeybindings(keybindings) {
  const list = assertArray(keybindings, "contributes.keybindings");
  const normalized = [];
  for (const [idx, kb] of list.entries()) {
    assertObject(kb, `contributes.keybindings[${idx}]`);
    const command = assertString(kb.command, `contributes.keybindings[${idx}].command`);
    const key = assertString(kb.key, `contributes.keybindings[${idx}].key`);
    const mac = assertOptionalString(kb.mac, `contributes.keybindings[${idx}].mac`);
    const when = assertOptionalString(kb.when, `contributes.keybindings[${idx}].when`);
    const normalizedKb = { ...kb, command, key };
    if (mac !== undefined) normalizedKb.mac = mac;
    if (when !== undefined) normalizedKb.when = when;
    normalized.push(normalizedKb);
  }
  return normalized;
}

function validateMenus(menus) {
  if (menus === undefined) return {};
  const obj = assertObject(menus, "contributes.menus");
  const normalizedMenus = {};
  for (const [menuId, items] of Object.entries(obj)) {
    const list = assertArray(items, `contributes.menus.${menuId}`);
    const normalizedItems = [];
    for (const [idx, item] of list.entries()) {
      assertObject(item, `contributes.menus.${menuId}[${idx}]`);
      const command = assertString(item.command, `contributes.menus.${menuId}[${idx}].command`);
      const when = assertOptionalString(item.when, `contributes.menus.${menuId}[${idx}].when`);
      const group = assertOptionalString(item.group, `contributes.menus.${menuId}[${idx}].group`);
      const normalizedItem = { ...item, command };
      if (when !== undefined) normalizedItem.when = when;
      if (group !== undefined) normalizedItem.group = group;
      normalizedItems.push(normalizedItem);
    }
    normalizedMenus[menuId] = normalizedItems;
  }
  return normalizedMenus;
}

function validateCustomFunctions(customFunctions) {
  const list = assertArray(customFunctions, "contributes.customFunctions");
  const seen = new Set();
  const normalized = [];
  for (const [idx, fn] of list.entries()) {
    assertObject(fn, `contributes.customFunctions[${idx}]`);
    const name = assertString(fn.name, `contributes.customFunctions[${idx}].name`);
    if (seen.has(name)) throw new ManifestError(`Duplicate custom function name: ${name}`);
    seen.add(name);
    const description = assertOptionalString(fn.description, `contributes.customFunctions[${idx}].description`);

    const params = assertArray(fn.parameters, `contributes.customFunctions[${idx}].parameters`);
    const normalizedParams = [];
    for (const [pIdx, p] of params.entries()) {
      assertObject(p, `contributes.customFunctions[${idx}].parameters[${pIdx}]`);
      const paramName = assertString(p.name, `contributes.customFunctions[${idx}].parameters[${pIdx}].name`);
      const paramType = assertString(p.type, `contributes.customFunctions[${idx}].parameters[${pIdx}].type`);
      const paramDescription = assertOptionalString(
        p.description,
        `contributes.customFunctions[${idx}].parameters[${pIdx}].description`,
      );
      const normalizedParam = { ...p, name: paramName, type: paramType };
      if (paramDescription !== undefined) normalizedParam.description = paramDescription;
      normalizedParams.push(normalizedParam);
    }

    assertObject(fn.result, `contributes.customFunctions[${idx}].result`);
    const resultType = assertString(fn.result.type, `contributes.customFunctions[${idx}].result.type`);
    const normalizedFn = { ...fn, name, parameters: normalizedParams, result: { ...fn.result, type: resultType } };
    if (description !== undefined) normalizedFn.description = description;
    normalized.push(normalizedFn);
  }
  return normalized;
}

function validateDataConnectors(dataConnectors) {
  const list = assertArray(dataConnectors, "contributes.dataConnectors");
  const seen = new Set();
  const normalized = [];
  for (const [idx, connector] of list.entries()) {
    assertObject(connector, `contributes.dataConnectors[${idx}]`);
    const id = assertString(connector.id, `contributes.dataConnectors[${idx}].id`);
    if (seen.has(id)) throw new ManifestError(`Duplicate dataConnector id: ${id}`);
    seen.add(id);
    const name = assertString(connector.name, `contributes.dataConnectors[${idx}].name`);
    const icon = assertOptionalString(connector.icon, `contributes.dataConnectors[${idx}].icon`);
    const normalizedConnector = { ...connector, id, name };
    if (icon !== undefined) normalizedConnector.icon = icon;
    normalized.push(normalizedConnector);
  }
  return normalized;
}

function validateConfiguration(configuration) {
  if (configuration === undefined) return undefined;
  const obj = assertObject(configuration, "contributes.configuration");
  const title = assertOptionalString(obj.title, "contributes.configuration.title");

  const properties = obj.properties ?? obj.settings;
  const props = assertObject(properties, "contributes.configuration.properties");

  const normalizedProps = {};
  for (const [key, prop] of Object.entries(props)) {
    assertObject(prop, `contributes.configuration.properties.${key}`);
    const type = assertString(prop.type, `contributes.configuration.properties.${key}.type`);
    const description = assertOptionalString(prop.description, `contributes.configuration.properties.${key}.description`);
    const normalizedProp = { ...prop, type };
    if (description !== undefined) normalizedProp.description = description;
    normalizedProps[key] = normalizedProp;
  }

  const normalizedConfiguration = { ...obj, properties: normalizedProps };
  if (title !== undefined) normalizedConfiguration.title = title;
  return normalizedConfiguration;
}

function validatePermissions(permissions) {
  const list = assertArray(permissions, "permissions");
  const normalized = [];
  for (const [idx, perm] of list.entries()) {
    if (typeof perm === "string") {
      const val = assertString(perm, `permissions[${idx}]`);
      if (!VALID_PERMISSIONS.has(val)) {
        throw new ManifestError(`Invalid permission: ${val}`);
      }
      normalized.push(val);
      continue;
    }

    if (perm && typeof perm === "object" && !Array.isArray(perm)) {
      const keys = Object.keys(perm);
      if (keys.length !== 1) {
        throw new ManifestError(
          `permissions[${idx}] must be a permission string or an object with a single permission key`
        );
      }
      const rawKey = keys[0];
      const key = String(rawKey).trim();
      if (key.length === 0) {
        throw new ManifestError(`Invalid permission: ${key}`);
      }
      if (!VALID_PERMISSIONS.has(key)) {
        throw new ManifestError(`Invalid permission: ${key}`);
      }
      normalized.push({ [key]: perm[rawKey] });
      continue;
    }

    throw new ManifestError(`permissions[${idx}] must be a permission string or object`);
  }
  return normalized;
}

function validateActivationEvents(activationEvents, contributes) {
  const list = assertArray(activationEvents, "activationEvents");
  const knownCommands = new Set((contributes.commands ?? []).map((c) => c.command));
  const knownPanels = new Set((contributes.panels ?? []).map((p) => p.id));
  const knownCustomFunctions = new Set((contributes.customFunctions ?? []).map((f) => f.name));
  const knownDataConnectors = new Set((contributes.dataConnectors ?? []).map((c) => c.id));

  const normalized = [];
  for (const [idx, event] of list.entries()) {
    const ev = assertString(event, `activationEvents[${idx}]`);
    if (ev === "onStartupFinished") {
      normalized.push(ev);
      continue;
    }

    if (ev.startsWith("onCommand:")) {
      const id = ev.slice("onCommand:".length).trim();
      if (!knownCommands.has(id)) {
        throw new ManifestError(`activationEvents references unknown command: ${id}`);
      }
      normalized.push(`onCommand:${id}`);
      continue;
    }

    if (ev.startsWith("onView:")) {
      const id = ev.slice("onView:".length).trim();
      if (id.length === 0) {
        throw new ManifestError("activationEvents references empty view/panel id");
      }
      // `onView:*` is used to activate an extension when its own contributed panel/view is opened.
      // Require that the referenced id is declared under `contributes.panels` to keep activation
      // events self-contained and validateable at publish time (marketplace).
      if (!knownPanels.has(id)) {
        throw new ManifestError(`activationEvents references unknown view/panel: ${id}`);
      }
      normalized.push(`onView:${id}`);
      continue;
    }

    if (ev.startsWith("onCustomFunction:")) {
      const name = ev.slice("onCustomFunction:".length).trim();
      if (!knownCustomFunctions.has(name)) {
        throw new ManifestError(`activationEvents references unknown custom function: ${name}`);
      }
      normalized.push(`onCustomFunction:${name}`);
      continue;
    }

    if (ev.startsWith("onDataConnector:")) {
      const id = ev.slice("onDataConnector:".length).trim();
      if (!knownDataConnectors.has(id)) {
        throw new ManifestError(`activationEvents references unknown data connector: ${id}`);
      }
      normalized.push(`onDataConnector:${id}`);
      continue;
    }

    throw new ManifestError(`Unsupported activation event: ${ev}`);
  }

  return normalized;
}

function validateExtensionManifest(manifest, options = {}) {
  const { engineVersion, enforceEngine = false } = options || {};

  const obj = assertObject(manifest, "manifest");

  const name = assertString(obj.name, "name");
  const version = assertString(obj.version, "version");
  const publisher = assertString(obj.publisher, "publisher");
  const main = assertString(obj.main, "main");
  const moduleEntry = assertOptionalString(obj.module, "module");
  const browserEntry = assertOptionalString(obj.browser, "browser");

  assertEntrypointExtension(main, "main", MAIN_ENTRYPOINT_EXTS);
  if (moduleEntry !== undefined) {
    if (moduleEntry.length === 0) {
      throw new ManifestError("module must be a non-empty string");
    }
    assertEntrypointExtension(moduleEntry, "module", ESM_ENTRYPOINT_EXTS);
  }
  if (browserEntry !== undefined) {
    if (browserEntry.length === 0) {
      throw new ManifestError("browser must be a non-empty string");
    }
    assertEntrypointExtension(browserEntry, "browser", ESM_ENTRYPOINT_EXTS);
  }

  if (!isValidSemver(version)) {
    throw new ManifestError(`Invalid version: ${version}`);
  }

  const engines = assertObject(obj.engines, "engines");
  const formulaRange = assertString(engines.formula, "engines.formula");

  if (enforceEngine) {
    if (typeof engineVersion !== "string" || engineVersion.trim().length === 0) {
      throw new ManifestError("engineVersion is required when enforceEngine is true");
    }

    if (!satisfies(engineVersion, formulaRange)) {
      throw new ManifestError(
        `Extension engine mismatch: formula ${engineVersion} does not satisfy ${formulaRange}`
      );
    }
  }

  const contributes = obj.contributes ? assertObject(obj.contributes, "contributes") : {};
  const commands = validateCommands(contributes.commands);
  const menus = validateMenus(contributes.menus);
  const keybindings = validateKeybindings(contributes.keybindings);
  const panels = validatePanels(contributes.panels);
  const customFunctions = validateCustomFunctions(contributes.customFunctions);
  const dataConnectors = validateDataConnectors(contributes.dataConnectors);
  const configuration = validateConfiguration(contributes.configuration);

  const validatedContributes = {
    commands,
    menus,
    keybindings,
    panels,
    customFunctions,
    dataConnectors,
    configuration,
  };

  const activationEvents = validateActivationEvents(obj.activationEvents, validatedContributes);
  const permissions = validatePermissions(obj.permissions);

  return {
    ...obj,
    name,
    version,
    publisher,
    main,
    engines: { ...engines, formula: formulaRange },
    contributes: validatedContributes,
    activationEvents,
    permissions,
    ...(moduleEntry !== undefined ? { module: moduleEntry } : {}),
    ...(browserEntry !== undefined ? { browser: browserEntry } : {}),
  };
}

const exportsObj = {
  ManifestError,
  VALID_PERMISSIONS,
  validateExtensionManifest,
};

// Make the validator usable from ESM in browser runtimes by stashing a copy of the exports on
// globalThis. The ESM wrapper (`index.mjs`) reads from this key to provide named exports without
// importing CommonJS (which browsers don't support).
try {
  if (typeof globalThis !== "undefined") {
    globalThis[GLOBAL_KEY] = exportsObj;
  }
} catch {
  // ignore
}

// CommonJS export (Node / marketplace).
try {
  if (typeof module !== "undefined" && module.exports) {
    module.exports = exportsObj;
  }
} catch {
  // ignore
}
