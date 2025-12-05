import { describe, it, expect } from "vitest";
import { extractCellReferences } from "../formulaParser";
import { createNameResolver } from "../nameResolver";
import type { DefinedNameInfo } from "../../types";

describe("extractCellReferences", () => {
  describe("basic cell references", () => {
    it("should extract simple cell reference", () => {
      const refs = extractCellReferences("=A1", "Sheet1");
      expect(refs).toEqual(["Sheet1!A1"]);
    });

    it("should extract multiple cell references", () => {
      const refs = extractCellReferences("=A1+B2", "Sheet1");
      expect(refs).toContain("Sheet1!A1");
      expect(refs).toContain("Sheet1!B2");
    });

    it("should handle absolute references", () => {
      const refs = extractCellReferences("=$A$1", "Sheet1");
      expect(refs).toEqual(["Sheet1!A1"]);
    });

    it("should handle mixed references", () => {
      const refs = extractCellReferences("=$A1+B$2", "Sheet1");
      expect(refs).toContain("Sheet1!A1");
      expect(refs).toContain("Sheet1!B2");
    });

    it("should return empty array for non-formula", () => {
      const refs = extractCellReferences("Hello", "Sheet1");
      expect(refs).toEqual([]);
    });

    it("should return empty array for empty string", () => {
      const refs = extractCellReferences("", "Sheet1");
      expect(refs).toEqual([]);
    });
  });

  describe("range references", () => {
    it("should expand range to individual cells", () => {
      const refs = extractCellReferences("=SUM(A1:A3)", "Sheet1");
      expect(refs).toContain("Sheet1!A1");
      expect(refs).toContain("Sheet1!A2");
      expect(refs).toContain("Sheet1!A3");
    });

    it("should expand 2D range", () => {
      const refs = extractCellReferences("=SUM(A1:B2)", "Sheet1");
      expect(refs).toContain("Sheet1!A1");
      expect(refs).toContain("Sheet1!A2");
      expect(refs).toContain("Sheet1!B1");
      expect(refs).toContain("Sheet1!B2");
    });
  });

  describe("cross-sheet references", () => {
    it("should extract cross-sheet reference", () => {
      const refs = extractCellReferences("=Sheet2!A1", "Sheet1");
      expect(refs).toEqual(["Sheet2!A1"]);
    });

    it("should handle quoted sheet names", () => {
      const refs = extractCellReferences("='My Sheet'!A1", "Sheet1");
      expect(refs).toEqual(["My Sheet!A1"]);
    });

    it("should extract cross-sheet range", () => {
      const refs = extractCellReferences("=SUM(Sheet2!A1:A2)", "Sheet1");
      expect(refs).toContain("Sheet2!A1");
      expect(refs).toContain("Sheet2!A2");
    });
  });

  describe("named references", () => {
    it("should resolve named reference to cell", () => {
      const definedNames: DefinedNameInfo[] = [
        { name: "mass", ref: "Sheet1!$A$1" },
      ];
      const nameResolver = createNameResolver(definedNames, ["Sheet1"]);

      const refs = extractCellReferences("=mass*2", "Sheet1", {
        nameResolver,
        currentSheetIndex: 0,
      });
      expect(refs).toContain("Sheet1!A1");
    });

    it("should resolve multiple named references", () => {
      const definedNames: DefinedNameInfo[] = [
        { name: "mass", ref: "Sheet1!$A$1" },
        { name: "velocity", ref: "Sheet1!$B$1" },
      ];
      const nameResolver = createNameResolver(definedNames, ["Sheet1"]);

      const refs = extractCellReferences("=mass*velocity", "Sheet1", {
        nameResolver,
        currentSheetIndex: 0,
      });
      expect(refs).toContain("Sheet1!A1");
      expect(refs).toContain("Sheet1!B1");
    });

    it("should resolve named range to multiple cells", () => {
      const definedNames: DefinedNameInfo[] = [
        { name: "myRange", ref: "Sheet1!$A$1:$A$3" },
      ];
      const nameResolver = createNameResolver(definedNames, ["Sheet1"]);

      const refs = extractCellReferences("=SUM(myRange)", "Sheet1", {
        nameResolver,
        currentSheetIndex: 0,
      });
      expect(refs).toContain("Sheet1!A1");
      expect(refs).toContain("Sheet1!A2");
      expect(refs).toContain("Sheet1!A3");
    });

    it("should resolve sheet-scoped name when on same sheet", () => {
      const definedNames: DefinedNameInfo[] = [
        { name: "localData", ref: "Sheet1!$C$1", sheetScope: 0 },
      ];
      const nameResolver = createNameResolver(definedNames, [
        "Sheet1",
        "Sheet2",
      ]);

      const refs = extractCellReferences("=localData+1", "Sheet1", {
        nameResolver,
        currentSheetIndex: 0,
      });
      expect(refs).toContain("Sheet1!C1");
    });

    it("should not resolve sheet-scoped name from different sheet", () => {
      const definedNames: DefinedNameInfo[] = [
        { name: "localData", ref: "Sheet1!$C$1", sheetScope: 0 },
      ];
      const nameResolver = createNameResolver(definedNames, [
        "Sheet1",
        "Sheet2",
      ]);

      // On Sheet2, localData (scoped to Sheet1) should not resolve
      const refs = extractCellReferences("=localData+1", "Sheet2", {
        nameResolver,
        currentSheetIndex: 1,
      });
      // Should not contain the cell since localData is not visible from Sheet2
      expect(refs).not.toContain("Sheet1!C1");
    });

    it("should not mistake function names for defined names", () => {
      const definedNames: DefinedNameInfo[] = [
        { name: "SUM", ref: "Sheet1!$A$1" }, // Someone named a cell SUM (bad idea but possible)
      ];
      const nameResolver = createNameResolver(definedNames, ["Sheet1"]);

      // SUM(B1:B3) should not resolve SUM as a name because it's followed by (
      const refs = extractCellReferences("=SUM(B1:B3)", "Sheet1", {
        nameResolver,
        currentSheetIndex: 0,
      });
      expect(refs).not.toContain("Sheet1!A1");
      expect(refs).toContain("Sheet1!B1");
    });

    it("should combine named and cell references", () => {
      const definedNames: DefinedNameInfo[] = [
        { name: "mass", ref: "Sheet1!$A$1" },
      ];
      const nameResolver = createNameResolver(definedNames, ["Sheet1"]);

      const refs = extractCellReferences("=mass+B1+C1", "Sheet1", {
        nameResolver,
        currentSheetIndex: 0,
      });
      expect(refs).toContain("Sheet1!A1"); // from mass
      expect(refs).toContain("Sheet1!B1"); // direct reference
      expect(refs).toContain("Sheet1!C1"); // direct reference
    });

    it("should handle case-insensitive name matching", () => {
      const definedNames: DefinedNameInfo[] = [
        { name: "MyData", ref: "Sheet1!$A$1" },
      ];
      const nameResolver = createNameResolver(definedNames, ["Sheet1"]);

      // Excel names are case-insensitive
      const refs = extractCellReferences("=mydata*2", "Sheet1", {
        nameResolver,
        currentSheetIndex: 0,
      });
      expect(refs).toContain("Sheet1!A1");
    });
  });
});
