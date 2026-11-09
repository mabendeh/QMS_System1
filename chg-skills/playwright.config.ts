import { defineConfig, devices } from "@playwright/test";

export default defineConfig({
  testDir: "tests/e2e",
  timeout: 60_000,
  workers: 1,
  fullyParallel: false,
  retries: 0,
  use: { baseURL: process.env.E2E_BASE_URL ?? "http://localhost:3100", trace: "off" },
  projects: [
    { name: "desktop", use: { ...devices["Desktop Chrome"], viewport: { width: 1440, height: 900 } } },
  ],
  webServer: process.env.E2E_BASE_URL
    ? undefined
    : {
        command: "bash tests/e2e/start-test-server.sh",
        url: "http://localhost:3100/login",
        timeout: 180_000,
        reuseExistingServer: false,
      },
});
