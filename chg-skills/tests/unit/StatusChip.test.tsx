import { describe, expect, it } from "vitest";
import { render, screen } from "@testing-library/react";
import { StatusChip, StatusLegend } from "@/components/StatusCell";

describe("StatusChip", () => {
  it("never relies on colour alone: shows an abbreviation and an accessible label", () => {
    render(<StatusChip status="COMPETENT" />);
    const chip = screen.getByRole("img", { name: /Competent/ });
    expect(chip).toHaveTextContent("C");
    expect(chip).toHaveAttribute("title", "Competent");
  });
  it("marks pending verification in text, not just style", () => {
    render(<StatusChip status="TRAINED" pending />);
    expect(screen.getByRole("img", { name: "Trained – pending verification" })).toBeInTheDocument();
  });
  it("legend explains every colour", () => {
    render(<StatusLegend />);
    for (const label of ["Needs Training", "In Training", "Trained", "Competent", "Expired", "Not Applicable / Not Assessed", "Pending verification"]) {
      expect(screen.getAllByText(label).length).toBeGreaterThan(0);
    }
  });
});
