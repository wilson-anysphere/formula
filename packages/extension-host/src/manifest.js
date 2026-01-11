const { isValidSemver, satisfies } = require("./semver");

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
  "ui.menus"
]);

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
  if (typeof value !== "string" || value.trim().length === 0) {
    throw new ManifestError(`${label} must be a non-empty string`);
  }
  return value;
}

function assertOptionalString(value, label) {
  if (value === undefined) return undefined;
  if (value === null) return undefined;
  if (typeof value !== "string") throw new ManifestError(`${label} must be a string`);
  return value;
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
  for (const [idx, cmd] of list.entries()) {
    assertObject(cmd, `contributes.commands[${idx}]`);
    const id = assertString(cmd.command, `contributes.commands[${idx}].command`);
    if (seen.has(id)) throw new ManifestError(`Duplicate command id: ${id}`);
    seen.add(id);
    assertString(cmd.title, `contributes.commands[${idx}].title`);
    assertOptionalString(cmd.category, `contributes.commands[${idx}].category`);
    assertOptionalString(cmd.icon, `contributes.commands[${idx}].icon`);
  }
  return list;
}

function validatePanels(panels) {
  const list = assertArray(panels, "contributes.panels");
  const seen = new Set();
  for (const [idx, panel] of list.entries()) {
    assertObject(panel, `contributes.panels[${idx}]`);
    const id = assertString(panel.id, `contributes.panels[${idx}].id`);
    if (seen.has(id)) throw new ManifestError(`Duplicate panel id: ${id}`);
    seen.add(id);
    assertString(panel.title, `contributes.panels[${idx}].title`);
    assertOptionalString(panel.icon, `contributes.panels[${idx}].icon`);
  }
  return list;
}

function validateKeybindings(keybindings) {
  const list = assertArray(keybindings, "contributes.keybindings");
  for (const [idx, kb] of list.entries()) {
    assertObject(kb, `contributes.keybindings[${idx}]`);
    assertString(kb.command, `contributes.keybindings[${idx}].command`);
    assertString(kb.key, `contributes.keybindings[${idx}].key`);
    assertOptionalString(kb.mac, `contributes.keybindings[${idx}].mac`);
  }
  return list;
}

function validateMenus(menus) {
  if (menus === undefined) return {};
  const obj = assertObject(menus, "contributes.menus");
  for (const [menuId, items] of Object.entries(obj)) {
    const list = assertArray(items, `contributes.menus.${menuId}`);
    for (const [idx, item] of list.entries()) {
      assertObject(item, `contributes.menus.${menuId}[${idx}]`);
      assertString(item.command, `contributes.menus.${menuId}[${idx}].command`);
      assertOptionalString(item.when, `contributes.menus.${menuId}[${idx}].when`);
      assertOptionalString(item.group, `contributes.menus.${menuId}[${idx}].group`);
    }
  }
  return obj;
}

function validateCustomFunctions(customFunctions) {
  const list = assertArray(customFunctions, "contributes.customFunctions");
  const seen = new Set();
  for (const [idx, fn] of list.entries()) {
    assertObject(fn, `contributes.customFunctions[${idx}]`);
    const name = assertString(fn.name, `contributes.customFunctions[${idx}].name`);
    if (seen.has(name)) throw new ManifestError(`Duplicate custom function name: ${name}`);
    seen.add(name);
    assertOptionalString(fn.description, `contributes.customFunctions[${idx}].description`);

    const params = assertArray(fn.parameters, `contributes.customFunctions[${idx}].parameters`);
    for (const [pIdx, p] of params.entries()) {
      assertObject(p, `contributes.customFunctions[${idx}].parameters[${pIdx}]`);
      assertString(p.name, `contributes.customFunctions[${idx}].parameters[${pIdx}].name`);
      assertString(p.type, `contributes.customFunctions[${idx}].parameters[${pIdx}].type`);
      assertOptionalString(
        p.description,
        `contributes.customFunctions[${idx}].parameters[${pIdx}].description`
      );
    }

    assertObject(fn.result, `contributes.customFunctions[${idx}].result`);
    assertString(fn.result.type, `contributes.customFunctions[${idx}].result.type`);
  }
  return list;
}

function validateDataConnectors(dataConnectors) {
  const list = assertArray(dataConnectors, "contributes.dataConnectors");
  const seen = new Set();
  for (const [idx, connector] of list.entries()) {
    assertObject(connector, `contributes.dataConnectors[${idx}]`);
    const id = assertString(connector.id, `contributes.dataConnectors[${idx}].id`);
    if (seen.has(id)) throw new ManifestError(`Duplicate dataConnector id: ${id}`);
    seen.add(id);
    assertString(connector.name, `contributes.dataConnectors[${idx}].name`);
    assertOptionalString(connector.icon, `contributes.dataConnectors[${idx}].icon`);
  }
  return list;
}

function validateConfiguration(configuration) {
  if (configuration === undefined) return undefined;
  const obj = assertObject(configuration, "contributes.configuration");
  assertOptionalString(obj.title, "contributes.configuration.title");

  const properties = obj.properties ?? obj.settings;
  const props = assertObject(properties, "contributes.configuration.properties");

  for (const [key, prop] of Object.entries(props)) {
    assertObject(prop, `contributes.configuration.properties.${key}`);
    assertString(prop.type, `contributes.configuration.properties.${key}.type`);
    assertOptionalString(prop.description, `contributes.configuration.properties.${key}.description`);
  }

  return { ...obj, properties: props };
}

function validatePermissions(permissions) {
  const list = assertArray(permissions, "permissions");
  for (const [idx, perm] of list.entries()) {
    const val = assertString(perm, `permissions[${idx}]`);
    if (!VALID_PERMISSIONS.has(val)) {
      throw new ManifestError(`Invalid permission: ${val}`);
    }
  }
  return list;
}

function validateActivationEvents(activationEvents, contributes) {
  const list = assertArray(activationEvents, "activationEvents");
  const knownCommands = new Set((contributes.commands ?? []).map((c) => c.command));
  const knownPanels = new Set((contributes.panels ?? []).map((p) => p.id));
  const knownCustomFunctions = new Set((contributes.customFunctions ?? []).map((f) => f.name));

  for (const [idx, event] of list.entries()) {
    const ev = assertString(event, `activationEvents[${idx}]`);
    if (ev === "onStartupFinished") continue;

    if (ev.startsWith("onCommand:")) {
      const id = ev.slice("onCommand:".length);
      if (!knownCommands.has(id)) {
        throw new ManifestError(`activationEvents references unknown command: ${id}`);
      }
      continue;
    }

    if (ev.startsWith("onView:")) {
      const id = ev.slice("onView:".length);
      if (!knownPanels.has(id)) {
        throw new ManifestError(`activationEvents references unknown view/panel: ${id}`);
      }
      continue;
    }

    if (ev.startsWith("onCustomFunction:")) {
      const name = ev.slice("onCustomFunction:".length);
      if (!knownCustomFunctions.has(name)) {
        throw new ManifestError(`activationEvents references unknown custom function: ${name}`);
      }
      continue;
    }

    throw new ManifestError(`Unsupported activation event: ${ev}`);
  }

  return list;
}

function validateExtensionManifest(manifest, { engineVersion }) {
  const obj = assertObject(manifest, "manifest");

  assertString(obj.name, "name");
  assertString(obj.version, "version");
  assertString(obj.publisher, "publisher");
  assertString(obj.main, "main");
  assertOptionalString(obj.module, "module");
  assertOptionalString(obj.browser, "browser");

  if (!isValidSemver(obj.version)) {
    throw new ManifestError(`Invalid version: ${obj.version}`);
  }

  const engines = assertObject(obj.engines, "engines");
  assertString(engines.formula, "engines.formula");

  if (!satisfies(engineVersion, engines.formula)) {
    throw new ManifestError(
      `Extension engine mismatch: formula ${engineVersion} does not satisfy ${engines.formula}`
    );
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
    configuration
  };

  validateActivationEvents(obj.activationEvents, validatedContributes);
  validatePermissions(obj.permissions);

  return {
    ...obj,
    contributes: validatedContributes
  };
}

module.exports = {
  ManifestError,
  VALID_PERMISSIONS,
  validateExtensionManifest
};
