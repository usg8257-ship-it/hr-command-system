# JavaScript Unit Test Generator

Generate Jest-compatible unit tests for JavaScript functions in this project.

## Instructions

When this skill is invoked:

1. **Identify the target** — if the user passes a filename or function name as `$ARGUMENTS`, focus on that. Otherwise ask which file or function to test.

2. **Read the source file** — understand the function signatures, inputs, outputs, and edge cases.

3. **Write tests** using the following conventions for this project:
   - Test framework: **Jest** (or plain QUnit-style if no framework is configured)
   - One `describe` block per function / agent
   - Cover: happy path, empty/null inputs, boundary values, error cases
   - Mock `google.script.run` calls using a simple stub pattern (see template below)

4. **Output the test file** at `__tests__/<source-filename>.test.js` (create the directory if needed).

5. **Summarise** what was generated and any untestable parts (e.g. direct DOM manipulation that requires a browser).

---

## GAS / Google Apps Script mock template

Use this stub at the top of every test file to avoid `google is not defined` errors:

```js
// Mock google.script.run
global.google = {
  script: {
    run: new Proxy({}, {
      get: (_, fn) => ({
        withSuccessHandler: (cb) => ({ withFailureHandler: () => ({ [fn]: (...args) => cb({success:true}) }) }),
        withFailureHandler: (cb) => ({ [fn]: () => {} })
      })
    })
  }
};
```

---

## Arguments

`$ARGUMENTS` — optional filename or function name to test (e.g. `script.html EmpAgent.save` or `Code.gs deleteEmployee`)
