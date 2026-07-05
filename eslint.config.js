import js from "@eslint/js";
import googleappsscript from "eslint-plugin-googleappsscript";

export default [
  js.configs.recommended, // Injects ESLint's base rules
  {
    files: ["**/*.js"],
    plugins: {
      googleappsscript: googleappsscript
    },
    languageOptions: {
      // Merges the Google Apps Script global variables into your environment
      globals: {
        ...googleappsscript.environments.googleappsscript.globals
      }
    },
    rules: {
      "no-unused-vars": "warn", // Warns about dead variables
      "no-undef": "error"       // Catches typo mistakes in functions/variables
    }
  }
];
