import { describe, it, expect } from "vitest";
import { createNameResolver, NameResolver } from "../nameResolver";
import type { DefinedNameInfo } from "../../types";

describe("NameResolver", () => {
  describe("createNameResolver", () => {
    it("should create a NameResolver instance", () => {
      const resolver = createNameResolver([], ["Sheet1"]);
      expect(resolver).toBeInstanceOf(NameResolver);
    });
  });

  describe("resolveName", () => {
    it("should resolve a global name to cell address", () => {
      const definedNames: DefinedNameInfo[] = [
        { name: "mass", ref: "Sheet1!$A$1" },
      ];
      const resolver = createNameResolver(definedNames, ["Sheet1"]);

      const cells = resolver.resolveName("mass", 0);
      expect(cells).toEqual(["Sheet1!A1"]);
    });

    it("should resolve a sheet-scoped name when on the same sheet", () => {
      const definedNames: DefinedNameInfo[] = [
        { name: "localVar", ref: "Sheet1!$B$2", sheetScope: 0 },
      ];
      const resolver = createNameResolver(definedNames, ["Sheet1", "Sheet2"]);

      const cells = resolver.resolveName("localVar", 0);
      expect(cells).toEqual(["Sheet1!B2"]);
    });

    it("should not resolve a sheet-scoped name from a different sheet", () => {
      const definedNames: DefinedNameInfo[] = [
        { name: "localVar", ref: "Sheet1!$B$2", sheetScope: 0 },
      ];
      const resolver = createNameResolver(definedNames, ["Sheet1", "Sheet2"]);

      const cells = resolver.resolveName("localVar", 1); // On Sheet2
      expect(cells).toEqual([]);
    });

    it("should prefer sheet-scoped name over global name on same sheet", () => {
      const definedNames: DefinedNameInfo[] = [
        { name: "myName", ref: "Sheet1!$A$1" }, // global
        { name: "myName", ref: "Sheet2!$B$2", sheetScope: 1 }, // scoped to Sheet2
      ];
      const resolver = createNameResolver(definedNames, ["Sheet1", "Sheet2"]);

      // On Sheet1, should get global
      expect(resolver.resolveName("myName", 0)).toEqual(["Sheet1!A1"]);

      // On Sheet2, should get scoped version
      expect(resolver.resolveName("myName", 1)).toEqual(["Sheet2!B2"]);
    });

    it("should resolve a range to multiple cells", () => {
      const definedNames: DefinedNameInfo[] = [
        { name: "myRange", ref: "Sheet1!$A$1:$B$2" },
      ];
      const resolver = createNameResolver(definedNames, ["Sheet1"]);

      const cells = resolver.resolveName("myRange", 0);
      expect(cells).toHaveLength(4);
      expect(cells).toContain("Sheet1!A1");
      expect(cells).toContain("Sheet1!A2");
      expect(cells).toContain("Sheet1!B1");
      expect(cells).toContain("Sheet1!B2");
    });

    it("should handle quoted sheet names", () => {
      const definedNames: DefinedNameInfo[] = [
        { name: "data", ref: "'My Sheet'!$C$3" },
      ];
      const resolver = createNameResolver(definedNames, ["My Sheet"]);

      const cells = resolver.resolveName("data", 0);
      expect(cells).toEqual(["My Sheet!C3"]);
    });

    it("should return empty array for unknown name", () => {
      const resolver = createNameResolver([], ["Sheet1"]);
      expect(resolver.resolveName("unknown", 0)).toEqual([]);
    });
  });

  describe("getDisplayName", () => {
    it("should return display name for a cell with global name", () => {
      const definedNames: DefinedNameInfo[] = [
        { name: "mass", ref: "Sheet1!$A$1" },
      ];
      const resolver = createNameResolver(definedNames, ["Sheet1"]);

      expect(resolver.getDisplayName("Sheet1!A1")).toBe("mass");
    });

    it("should return undefined for cell without name", () => {
      const definedNames: DefinedNameInfo[] = [
        { name: "mass", ref: "Sheet1!$A$1" },
      ];
      const resolver = createNameResolver(definedNames, ["Sheet1"]);

      expect(resolver.getDisplayName("Sheet1!B2")).toBeUndefined();
    });

    it("should not return display name for ranges (only single cells)", () => {
      const definedNames: DefinedNameInfo[] = [
        { name: "myRange", ref: "Sheet1!$A$1:$B$2" },
      ];
      const resolver = createNameResolver(definedNames, ["Sheet1"]);

      // Ranges don't get display names (would be confusing)
      expect(resolver.getDisplayName("Sheet1!A1")).toBeUndefined();
    });

    it("should show scoped display name when ambiguous", () => {
      const definedNames: DefinedNameInfo[] = [
        { name: "data", ref: "Sheet1!$A$1", sheetScope: 0 },
        { name: "data", ref: "Sheet2!$A$1", sheetScope: 1 },
      ];
      const resolver = createNameResolver(definedNames, ["Sheet1", "Sheet2"]);

      // When same name exists in multiple scopes, show "SheetName::name"
      expect(resolver.getDisplayName("Sheet1!A1")).toBe("Sheet1::data");
      expect(resolver.getDisplayName("Sheet2!A1")).toBe("Sheet2::data");
    });
  });

  describe("getAllNames", () => {
    it("should return all defined names sorted by length descending", () => {
      const definedNames: DefinedNameInfo[] = [
        { name: "a", ref: "Sheet1!$A$1" },
        { name: "abc", ref: "Sheet1!$B$1" },
        { name: "ab", ref: "Sheet1!$C$1" },
      ];
      const resolver = createNameResolver(definedNames, ["Sheet1"]);

      const names = resolver.getAllNames();
      expect(names).toEqual(["abc", "ab", "a"]);
    });

    it("should deduplicate names that exist at multiple scopes", () => {
      const definedNames: DefinedNameInfo[] = [
        { name: "data", ref: "Sheet1!$A$1" }, // global
        { name: "data", ref: "Sheet2!$A$1", sheetScope: 1 }, // scoped
      ];
      const resolver = createNameResolver(definedNames, ["Sheet1", "Sheet2"]);

      const names = resolver.getAllNames();
      expect(names).toEqual(["data"]);
    });
  });

  describe("isDefinedName", () => {
    it("should return true for defined names", () => {
      const definedNames: DefinedNameInfo[] = [
        { name: "mass", ref: "Sheet1!$A$1" },
      ];
      const resolver = createNameResolver(definedNames, ["Sheet1"]);

      expect(resolver.isDefinedName("mass")).toBe(true);
    });

    it("should return false for undefined names", () => {
      const definedNames: DefinedNameInfo[] = [
        { name: "mass", ref: "Sheet1!$A$1" },
      ];
      const resolver = createNameResolver(definedNames, ["Sheet1"]);

      expect(resolver.isDefinedName("unknown")).toBe(false);
    });
  });
});
