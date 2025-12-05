import { describe, it, expect, beforeEach } from "vitest";
import {
  createEmptyMetadata,
  loadMetadataFromStorage,
  saveMetadataToStorage,
  clearMetadataFromStorage,
  expandCellKey,
  findExistingName,
  findNotesForCell,
  getNameForCell,
  validateNameAssignment,
  setName,
  removeName,
  addNote,
  updateNote,
  removeNote,
  compareMetadata,
  buildCellNameMap,
} from "../metadataStore";
import type { CellMetadata } from "../../types";

describe("metadataStore", () => {
  describe("createEmptyMetadata", () => {
    it("should create empty metadata with no names or notes", () => {
      const metadata = createEmptyMetadata();
      expect(metadata.names).toEqual([]);
      expect(metadata.notes).toEqual([]);
    });
  });

  describe("localStorage operations", () => {
    const testFileName = "test-file.xlsx";

    beforeEach(() => {
      localStorage.clear();
    });

    it("should return null for non-existent metadata", () => {
      const result = loadMetadataFromStorage("non-existent.xlsx");
      expect(result).toBeNull();
    });

    it("should save and load metadata correctly", () => {
      const metadata: CellMetadata = {
        names: [{ cellKey: "Sheet1!A1", name: "TestName" }],
        notes: [{ cellKey: "Sheet1!A1", note: "TestNote" }],
      };

      saveMetadataToStorage(testFileName, metadata);
      const loaded = loadMetadataFromStorage(testFileName);

      expect(loaded).toEqual(metadata);
    });

    it("should clear metadata correctly", () => {
      const metadata: CellMetadata = {
        names: [{ cellKey: "Sheet1!A1", name: "TestName" }],
        notes: [],
      };

      saveMetadataToStorage(testFileName, metadata);
      clearMetadataFromStorage(testFileName);

      const loaded = loadMetadataFromStorage(testFileName);
      expect(loaded).toBeNull();
    });
  });

  describe("expandCellKey", () => {
    it("should handle single cell reference", () => {
      const result = expandCellKey("Sheet1!A1");
      expect(result).toEqual(["Sheet1!A1"]);
    });

    it("should handle comma-separated cells", () => {
      const result = expandCellKey("Sheet1!A1,Sheet1!B2");
      expect(result).toEqual(["Sheet1!A1", "Sheet1!B2"]);
    });

    it("should handle cell range", () => {
      const result = expandCellKey("Sheet1!A1:B2");
      expect(result).toEqual([
        "Sheet1!A1",
        "Sheet1!B1",
        "Sheet1!A2",
        "Sheet1!B2",
      ]);
    });

    it("should handle range with column letters", () => {
      const result = expandCellKey("Sheet1!A1:A3");
      expect(result).toEqual(["Sheet1!A1", "Sheet1!A2", "Sheet1!A3"]);
    });

    it("should handle quoted sheet names", () => {
      const result = expandCellKey("'My Sheet'!A1");
      expect(result).toEqual(["My Sheet!A1"]);
    });

    it("should uppercase cell references", () => {
      const result = expandCellKey("Sheet1!a1");
      expect(result).toEqual(["Sheet1!A1"]);
    });

    it("should return empty array for invalid reference", () => {
      const result = expandCellKey("invalid");
      expect(result).toEqual([]);
    });
  });

  describe("Name Management", () => {
    describe("setName", () => {
      it("should add a new name entry", () => {
        const metadata = createEmptyMetadata();
        const result = setName(metadata, "Sheet1!A1", "MyName");

        expect(result.names).toHaveLength(1);
        expect(result.names[0]).toEqual({
          cellKey: "Sheet1!A1",
          name: "MyName",
        });
      });

      it("should update existing name for same cellKey", () => {
        let metadata = createEmptyMetadata();
        metadata = setName(metadata, "Sheet1!A1", "OldName");
        metadata = setName(metadata, "Sheet1!A1", "NewName");

        expect(metadata.names).toHaveLength(1);
        expect(metadata.names[0].name).toBe("NewName");
      });

      it("should preserve other entries when updating", () => {
        let metadata = createEmptyMetadata();
        metadata = setName(metadata, "Sheet1!A1", "Name1");
        metadata = setName(metadata, "Sheet1!B1", "Name2");
        metadata = setName(metadata, "Sheet1!A1", "Name1Updated");

        expect(metadata.names).toHaveLength(2);
        expect(
          metadata.names.find((n) => n.cellKey === "Sheet1!A1")?.name
        ).toBe("Name1Updated");
        expect(
          metadata.names.find((n) => n.cellKey === "Sheet1!B1")?.name
        ).toBe("Name2");
      });
    });

    describe("removeName", () => {
      it("should remove a name entry by cellKey", () => {
        let metadata = createEmptyMetadata();
        metadata = setName(metadata, "Sheet1!A1", "MyName");
        metadata = removeName(metadata, "Sheet1!A1");

        expect(metadata.names).toHaveLength(0);
      });

      it("should not affect other names", () => {
        let metadata = createEmptyMetadata();
        metadata = setName(metadata, "Sheet1!A1", "Name1");
        metadata = setName(metadata, "Sheet1!B1", "Name2");
        metadata = removeName(metadata, "Sheet1!A1");

        expect(metadata.names).toHaveLength(1);
        expect(metadata.names[0].name).toBe("Name2");
      });
    });

    describe("getNameForCell", () => {
      it("should return name for a cell", () => {
        let metadata = createEmptyMetadata();
        metadata = setName(metadata, "Sheet1!A1", "MyName");

        expect(getNameForCell("Sheet1!A1", metadata)).toBe("MyName");
      });

      it("should return null for unnamed cell", () => {
        const metadata = createEmptyMetadata();
        expect(getNameForCell("Sheet1!A1", metadata)).toBeNull();
      });

      it("should find name via range expansion", () => {
        let metadata = createEmptyMetadata();
        metadata = setName(metadata, "Sheet1!A1:A3", "RangeName");

        expect(getNameForCell("Sheet1!A2", metadata)).toBe("RangeName");
      });
    });

    describe("findExistingName", () => {
      it("should find name entry for direct match", () => {
        let metadata = createEmptyMetadata();
        metadata = setName(metadata, "Sheet1!A1", "MyName");

        const entry = findExistingName("Sheet1!A1", metadata);
        expect(entry).toEqual({ cellKey: "Sheet1!A1", name: "MyName" });
      });

      it("should find name entry via range", () => {
        let metadata = createEmptyMetadata();
        metadata = setName(metadata, "Sheet1!A1:B2", "RangeName");

        const entry = findExistingName("Sheet1!B2", metadata);
        expect(entry).toEqual({ cellKey: "Sheet1!A1:B2", name: "RangeName" });
      });

      it("should return null when no name exists", () => {
        const metadata = createEmptyMetadata();
        expect(findExistingName("Sheet1!A1", metadata)).toBeNull();
      });
    });
  });

  describe("Name Uniqueness Validation", () => {
    describe("validateNameAssignment", () => {
      it("should return empty array for valid assignment", () => {
        const metadata = createEmptyMetadata();
        const conflicts = validateNameAssignment("Sheet1!A1", metadata);
        expect(conflicts).toEqual([]);
      });

      it("should return conflict when cell already has a name", () => {
        let metadata = createEmptyMetadata();
        metadata = setName(metadata, "Sheet1!A1", "ExistingName");

        const conflicts = validateNameAssignment("Sheet1!A1", metadata);
        expect(conflicts).toHaveLength(1);
        expect(conflicts[0]).toContain("Sheet1!A1");
        expect(conflicts[0]).toContain("ExistingName");
      });

      it("should return conflict when range overlaps with named cell", () => {
        let metadata = createEmptyMetadata();
        metadata = setName(metadata, "Sheet1!B2", "ExistingName");

        // Try to name a range that includes B2
        const conflicts = validateNameAssignment("Sheet1!A1:C3", metadata);
        expect(conflicts).toHaveLength(1);
        expect(conflicts[0]).toContain("Sheet1!B2");
      });

      it("should return conflict when single cell overlaps with named range", () => {
        let metadata = createEmptyMetadata();
        metadata = setName(metadata, "Sheet1!A1:C3", "RangeName");

        // Try to name a cell within the range
        const conflicts = validateNameAssignment("Sheet1!B2", metadata);
        expect(conflicts).toHaveLength(1);
        expect(conflicts[0]).toContain("RangeName");
      });

      it("should exclude current entry when editing", () => {
        let metadata = createEmptyMetadata();
        metadata = setName(metadata, "Sheet1!A1", "MyName");

        // When editing, should not conflict with itself
        const conflicts = validateNameAssignment(
          "Sheet1!A1",
          metadata,
          "Sheet1!A1"
        );
        expect(conflicts).toEqual([]);
      });

      it("should still detect conflicts with other entries when editing", () => {
        let metadata = createEmptyMetadata();
        metadata = setName(metadata, "Sheet1!A1", "Name1");
        metadata = setName(metadata, "Sheet1!B1", "Name2");

        // Trying to expand A1's range to include B1 (which has Name2)
        const conflicts = validateNameAssignment(
          "Sheet1!A1:B1",
          metadata,
          "Sheet1!A1"
        );
        expect(conflicts).toHaveLength(1);
        expect(conflicts[0]).toContain("Name2");
      });
    });
  });

  describe("Note Management", () => {
    describe("addNote", () => {
      it("should add a note entry", () => {
        const metadata = createEmptyMetadata();
        const result = addNote(metadata, "Sheet1!A1", "My note");

        expect(result.notes).toHaveLength(1);
        expect(result.notes[0]).toEqual({
          cellKey: "Sheet1!A1",
          note: "My note",
        });
      });

      it("should allow multiple notes on same cell", () => {
        let metadata = createEmptyMetadata();
        metadata = addNote(metadata, "Sheet1!A1", "Note 1");
        metadata = addNote(metadata, "Sheet1!A1", "Note 2");
        metadata = addNote(metadata, "Sheet1!A1", "Note 3");

        expect(metadata.notes).toHaveLength(3);
        expect(metadata.notes.map((n) => n.note)).toEqual([
          "Note 1",
          "Note 2",
          "Note 3",
        ]);
      });
    });

    describe("updateNote", () => {
      it("should update an existing note", () => {
        let metadata = createEmptyMetadata();
        metadata = addNote(metadata, "Sheet1!A1", "Old note");
        metadata = updateNote(metadata, "Sheet1!A1", "Old note", "New note");

        expect(metadata.notes).toHaveLength(1);
        expect(metadata.notes[0].note).toBe("New note");
      });

      it("should only update the matching note", () => {
        let metadata = createEmptyMetadata();
        metadata = addNote(metadata, "Sheet1!A1", "Note 1");
        metadata = addNote(metadata, "Sheet1!A1", "Note 2");
        metadata = updateNote(
          metadata,
          "Sheet1!A1",
          "Note 1",
          "Updated Note 1"
        );

        expect(metadata.notes).toHaveLength(2);
        expect(metadata.notes[0].note).toBe("Updated Note 1");
        expect(metadata.notes[1].note).toBe("Note 2");
      });
    });

    describe("removeNote", () => {
      it("should remove a specific note", () => {
        let metadata = createEmptyMetadata();
        metadata = addNote(metadata, "Sheet1!A1", "Note 1");
        metadata = addNote(metadata, "Sheet1!A1", "Note 2");
        metadata = removeNote(metadata, "Sheet1!A1", "Note 1");

        expect(metadata.notes).toHaveLength(1);
        expect(metadata.notes[0].note).toBe("Note 2");
      });
    });

    describe("findNotesForCell", () => {
      it("should find all notes for a cell", () => {
        let metadata = createEmptyMetadata();
        metadata = addNote(metadata, "Sheet1!A1", "Note 1");
        metadata = addNote(metadata, "Sheet1!A1", "Note 2");
        metadata = addNote(metadata, "Sheet1!B1", "Other note");

        const notes = findNotesForCell("Sheet1!A1", metadata);
        expect(notes).toHaveLength(2);
        expect(notes.map((n) => n.note)).toEqual(["Note 1", "Note 2"]);
      });

      it("should find notes from ranges containing the cell", () => {
        let metadata = createEmptyMetadata();
        metadata = addNote(metadata, "Sheet1!A1:B2", "Range note");

        const notes = findNotesForCell("Sheet1!B2", metadata);
        expect(notes).toHaveLength(1);
        expect(notes[0].note).toBe("Range note");
      });

      it("should return empty array when no notes exist", () => {
        const metadata = createEmptyMetadata();
        const notes = findNotesForCell("Sheet1!A1", metadata);
        expect(notes).toEqual([]);
      });
    });
  });

  describe("Metadata Comparison", () => {
    describe("compareMetadata", () => {
      it("should return no conflict when both are null", () => {
        const result = compareMetadata(null, null);
        expect(result.hasConflict).toBe(false);
        expect(result.differences).toEqual([]);
      });

      it("should return no conflict when identical", () => {
        const metadata: CellMetadata = {
          names: [{ cellKey: "Sheet1!A1", name: "Name1" }],
          notes: [{ cellKey: "Sheet1!A1", note: "Note1" }],
        };

        const result = compareMetadata(metadata, metadata);
        expect(result.hasConflict).toBe(false);
      });

      it("should detect only-local name differences", () => {
        const local: CellMetadata = {
          names: [{ cellKey: "Sheet1!A1", name: "LocalName" }],
          notes: [],
        };
        const excel: CellMetadata = { names: [], notes: [] };

        const result = compareMetadata(local, excel);
        expect(result.hasConflict).toBe(true);
        expect(result.differences).toHaveLength(1);
        expect(result.differences[0].status).toBe("only-local");
        expect(result.differences[0].type).toBe("name");
      });

      it("should detect only-excel name differences", () => {
        const local: CellMetadata = { names: [], notes: [] };
        const excel: CellMetadata = {
          names: [{ cellKey: "Sheet1!A1", name: "ExcelName" }],
          notes: [],
        };

        const result = compareMetadata(local, excel);
        expect(result.hasConflict).toBe(true);
        expect(result.differences).toHaveLength(1);
        expect(result.differences[0].status).toBe("only-excel");
      });

      it("should detect different values for same key", () => {
        const local: CellMetadata = {
          names: [{ cellKey: "Sheet1!A1", name: "LocalName" }],
          notes: [],
        };
        const excel: CellMetadata = {
          names: [{ cellKey: "Sheet1!A1", name: "ExcelName" }],
          notes: [],
        };

        const result = compareMetadata(local, excel);
        expect(result.hasConflict).toBe(true);
        expect(result.differences).toHaveLength(1);
        expect(result.differences[0].status).toBe("different");
        expect(result.differences[0].localValue).toBe("LocalName");
        expect(result.differences[0].excelValue).toBe("ExcelName");
      });

      it("should detect note differences", () => {
        const local: CellMetadata = {
          names: [],
          notes: [{ cellKey: "Sheet1!A1", note: "Local note" }],
        };
        const excel: CellMetadata = { names: [], notes: [] };

        const result = compareMetadata(local, excel);
        expect(result.hasConflict).toBe(true);
        expect(result.differences[0].type).toBe("note");
      });
    });
  });

  describe("buildCellNameMap", () => {
    it("should create correct mapping for single cells", () => {
      let metadata = createEmptyMetadata();
      metadata = setName(metadata, "Sheet1!A1", "Name1");
      metadata = setName(metadata, "Sheet1!B1", "Name2");

      const map = buildCellNameMap(metadata);

      expect(map.get("Sheet1!A1")).toBe("Name1");
      expect(map.get("Sheet1!B1")).toBe("Name2");
      expect(map.get("Sheet1!C1")).toBeUndefined();
    });

    it("should expand ranges into individual cell mappings", () => {
      let metadata = createEmptyMetadata();
      metadata = setName(metadata, "Sheet1!A1:A3", "RangeName");

      const map = buildCellNameMap(metadata);

      expect(map.get("Sheet1!A1")).toBe("RangeName");
      expect(map.get("Sheet1!A2")).toBe("RangeName");
      expect(map.get("Sheet1!A3")).toBe("RangeName");
      expect(map.get("Sheet1!A4")).toBeUndefined();
    });
  });
});
