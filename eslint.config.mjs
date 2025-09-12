import eslint from "@eslint/js";
import tseslint from "typescript-eslint";
import playwright from "eslint-plugin-playwright";
import prettier from "eslint-config-prettier";
import { fileURLToPath } from "node:url";

const tsconfigRootDir = fileURLToPath(new URL(".", import.meta.url));

const typedConfigs = tseslint.configs.recommendedTypeChecked.map((cfg) => ({
  ...cfg,
  files: ["**/*.ts", "**/*.tsx"],
  languageOptions: {
    ...cfg.languageOptions,
    parser: tseslint.parser,
    parserOptions: {
      ...cfg.languageOptions?.parserOptions,
      projectService: true,
      project: ["./tsconfig.json"],
      tsconfigRootDir,
    },
  },
}));

export default [
  eslint.configs.recommended,
  ...tseslint.configs.recommended,
  ...typedConfigs,
  playwright.configs["flat/recommended"],
  // Disable formatting rules conflicting with Prettier
  prettier,
  {
    files: ["**/*.ts", "**/*.tsx"],
    rules: {
      "@typescript-eslint/no-floating-promises": "error",
      "@typescript-eslint/no-misused-promises": "error",
      "@typescript-eslint/await-thenable": "error",
    },
  },
];
