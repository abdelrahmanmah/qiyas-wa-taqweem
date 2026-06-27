import { compileSchema } from "./compileSchema.js";

// Vite resolves this glob at build time. The yaml-loader plugin transforms
// each .yaml file into a JS module that exports the parsed object.
const rawModules = import.meta.glob("./*.yaml", { eager: true });

export const SCHEMAS = Object.fromEntries(
  Object.values(rawModules).map((m) => {
    const schema = compileSchema(m.default);
    return [schema.id, schema];
  })
);
