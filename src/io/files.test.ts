import { describe, expect, it } from "vitest";
import { mapInputFiles, missingRequiredInputs } from "./files";

function makeFile(name: string): File {
  return new File([""], name);
}

describe("input file mapping", () => {
  it("requires the two private templates for parity output", () => {
    expect(missingRequiredInputs({})).toContain("Nagwa Technologies.xlsx template");
    expect(missingRequiredInputs({})).toContain("Final Nagwa Technologies.xlsx template");
  });

  it("maps the expected report and template basenames", () => {
    const mapped = mapInputFiles([
      makeFile("Attendance Report.xls"),
      makeFile("Nagwa Technologies.xlsx"),
      makeFile("Final Nagwa Technologies.xlsx"),
    ]);

    expect(mapped.attendance?.name).toBe("Attendance Report.xls");
    expect(mapped.nagwaTemplate?.name).toBe("Nagwa Technologies.xlsx");
    expect(mapped.finalTemplate?.name).toBe("Final Nagwa Technologies.xlsx");
  });
});
