# Playwright E2E: Excel Online TODAY() Function

End‑to‑end automated test (TypeScript + Playwright) that verifies Excel Online correctly evaluates the standard Excel function `=TODAY()` in a new workbook (Chrome / Chromium).

## Goal

1. Launch browser
2. Log into Excel Online with predefined Microsoft account credentials
3. Create a new blank workbook
4. Enter `=TODAY()` in cell `A2`
5. Format the cell as a date (configured format)
6. Assert the displayed value equals the current date for the test run

## Stack

- Playwright Test (`@playwright/test`)
- TypeScript
- Page Object Model (POM)
- ESLint (flat config, type‑aware) + Prettier
- Moment.js (date formatting consistency)

## Repository Structure (key parts)

```
src/
  pages/                # Page objects (login, launch Excel, workbook interactions)
  tests/                # Test specs (excel-online.spec.ts)
  types/                # Shared TypeScript types (cell parameters, formatting options)
eslint.config.mjs       # Flat ESLint configuration (ESLint v9)
playwright.config.ts    # Playwright configuration (can be extended for video/trace)
tsconfig.json           # TypeScript compiler options
.prettierrc.json        # Prettier formatting rules
.vscode/settings.json   # On-save format + organize imports + lint fixes
```

## Prerequisites

- Node.js 18+ (LTS recommended)
- A Microsoft 365 / Outlook account with access to Excel Online

## Credential & Config Management

Environment variables are used (NOT hard‑coded). Create a `.env` file (NOT committed) in the project root:

```
USER=your_email@domain.com
PASSWORD=yourStrongPassword
```

Test user that was used during implementation (Configured to login with password)

```
USER=zabrinayellow@powerscrews.com
PASSWORD=HasloBardzoTajne
```

Playwright tests reference `process.env.USER` / `process.env.PASSWORD` directly.

## Install

```
npm install
```

## Core Scripts

```
npm test            # Run all tests (headless by default)
npm run test:headed # Run with UI
npm run test:debug  # Debug mode (Playwright inspector)
npm run test:report # Open last HTML report

npm run pretest     # Lint + Prettier check + type check (auto before test)
npm run lint:fix    # ESLint auto-fixes
npm run format      # Apply Prettier formatting
npm run format:check# Check formatting only
```

Run a single test:

```
npx playwright test src/tests/excel-online.spec.ts
```

Filter by title:

```
npx playwright test -g "function: today"
```

## Test Flow Details

1. Login Page Object performs authentication (email + password)
2. Launch Page clicks "Create a new blank workbook"
3. Workbook Page selects cell A2, types formula, reselects, opens the Format pane, applies a date format
4. Polls until retrieved cell metadata (parsed from accessibility attributes) matches expected
5. Date comparison uses a consistent format string (e.g. `MM.DD.YY`) to avoid locale ambiguity

## Date Handling & Timezone Notes

- Excel Online may evaluate `TODAY()` in a server or tenant timezone (often UTC or tenant region)
- If you experience off‑by‑one day failures around midnight boundaries, consider:
  - Normalizing to UTC
  - Allowing ±1 day tolerance
  - Running the test away from midnight in the target timezone

## Formatting

On save, VS Code (workspace settings) automatically:

- Formats with Prettier
- Organizes imports
- Applies ESLint source fixes
- Auto-saves after 5s idle (configurable)

## Extending / Modifying Selectors

If Microsoft updates ARIA roles / labels:

- Use Playwright trace viewer (`npx playwright show-trace`) or DevTools to inspect new attributes
- Update the corresponding page object only (tests stay stable)

## Known Limitations / Bottlenecks

- Relies on live Microsoft service (network latency + possible flakiness)
- MFA-enabled accounts require app passwords or disabled MFA for test user
- UI selector brittleness if Excel Online layout changes
- Running in parallel against the same account could cause session conflicts

## Possible Improvements

- Add retry logic or soft assertions around initial workbook load
- Capture video & trace automatically (enable in `playwright.config.ts`)
- Add GitHub Actions CI pipeline (cache browsers, store traces as artifacts)

## Enabling Trace / Video (Optional)

In `playwright.config.ts`:

```ts
use: {
  trace: 'on-first-retry',
  video: 'retain-on-failure'
}
```

## Sample GitHub Actions Workflow (Optional)

```yml
name: CI
on: [push, pull_request]
jobs:
  tests:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-node@v4
        with:
          node-version: 20
      - run: npm ci
      - run: npx playwright install --with-deps
      - run: npm test
        env:
          USER: ${{ secrets.EXCEL_USER }}
          PASSWORD: ${{ secrets.EXCEL_PASSWORD }}
      - uses: actions/upload-artifact@v4
        if: always()
        with:
          name: playwright-report
          path: playwright-report
```

## FAQ

**Q: Where do I put credentials?** `.env` file (excluded from VCS) or CI secrets.

**Q: Test fails with wrong date (one day ahead/behind).** Check timezone; switch to UTC formatting or add ±1 day tolerance.

**Q: Why Moment.js?** Simplicity.

**Q: How do I add another formula test?** Create a new spec or parametrize the existing one; reuse page objects.

**Q: How to regenerate HTML report?** `npm test` then `npm run test:report`.

**Q: Can I run only this test in headed mode?** `npx playwright test -g today --headed`.

**Q: How to debug selectors?** Use `PWDEBUG=1 npx playwright test` or `test:debug` script.

**Q: Videos / traces not captured?** Enable in config (see above) and re-run failing test.

## License

MIT (see `package.json`).

---

Feel free to open issues / extend with more Excel function validations.
