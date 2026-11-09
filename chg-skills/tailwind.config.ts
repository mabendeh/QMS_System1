import type { Config } from "tailwindcss";

export default {
  content: ["./src/**/*.{ts,tsx}"],
  theme: {
    extend: {
      colors: {
        navy: { 50: "#EEF2F8", 100: "#D7E0EE", 200: "#AFC1DC", 600: "#24427A", 700: "#1B3463", 800: "#14284B", 900: "#0E1C36" },
        ink: "#1A2433",
        muted: "#5D6B7E",
        line: "#D9DFE7",
        canvas: "#F4F6F9",
        st: {
          red: "#C62828", yellow: "#F2C230", green: "#2E8B57", blue: "#1E63C4", gray: "#E3E7EC",
        },
      },
      fontFamily: { sans: ["var(--font-sans)", "Segoe UI", "Arial", "sans-serif"] },
    },
  },
  plugins: [],
} satisfies Config;
