import pluginJs from "@eslint/js";
import googleappsscript from "eslint-plugin-googleappsscript";
import globals from "globals";

export default [
  pluginJs.configs.recommended, 

  {
    files: ["**/*.js"],
    plugins: {
      googleappsscript: googleappsscript
    },
    languageOptions: {
      ecmaVersion: "latest",
      sourceType: "module", 
      globals: {
        ...globals.browser,
        ...googleappsscript.environments.googleappsscript.globals
      }
    },
    rules: {
      // turn off both to remove warnings since Apps Script merges files at runtime
      "no-unused-vars": "off", 
      "no-undef": "off"
    }
  }
];
