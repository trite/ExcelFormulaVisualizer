/**
 * Script to generate test Excel files for the Excel Formula Visualizer
 * Run with: npx tsx scripts/generate-test-files.ts
 */

import * as XLSX from "xlsx";
import * as path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const FIXTURES_DIR = path.join(__dirname, "..", "tests", "fixtures");

/**
 * Test File 1: Simple single-sheet with physics/math formulas
 *
 * Cluster 1 (Pythagorean theorem + Force calculations):
 * - A1: side_a = 3
 * - A2: side_b = 4
 * - A3: hypotenuse = SQRT(A1^2 + A2^2) = 5
 * - B1: mass = 10
 * - B2: acceleration = 9.8
 * - B3: force = B1 * B2 (F = ma)
 * - C1: velocity = 20
 * - C2: time = 5
 * - C3: distance = C1 * C2 (d = vt)
 * - D1: total = A3 + B3 + C3 (combines all three)
 *
 * Cluster 2 (isolated):
 * - E1: x = 100
 * - E2: y = E1 * 2
 */
function createSimpleTestFile(): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  const wsData = [
    // Row 1: base values
    { A: 3, B: 10, C: 20, D: null, E: 100 },
    // Row 2: more base values
    { A: 4, B: 9.8, C: 5, D: null, E: null },
    // Row 3: calculated values
    { A: null, B: null, C: null, D: null, E: null },
  ];

  const ws = XLSX.utils.aoa_to_sheet([
    [3, 10, 20, null, 100], // Row 1
    [4, 9.8, 5, null, null], // Row 2
    [null, null, null, null, null], // Row 3
  ]);

  // Add formulas
  // Pythagorean: hypotenuse = sqrt(a² + b²)
  ws["A3"] = { t: "n", f: "SQRT(A1*A1+A2*A2)", v: 5 };

  // Force = mass * acceleration
  ws["B3"] = { t: "n", f: "B1*B2", v: 98 };

  // Distance = velocity * time
  ws["C3"] = { t: "n", f: "C1*C2", v: 100 };

  // Total combining all calculations
  ws["D1"] = { t: "n", f: "A3+B3+C3", v: 203 };

  // Isolated cluster
  ws["E2"] = { t: "n", f: "E1*2", v: 200 };

  // Add labels in column F for clarity
  ws["F1"] = { t: "s", v: "side_a, mass, velocity" };
  ws["F2"] = { t: "s", v: "side_b, accel, time" };
  ws["F3"] = { t: "s", v: "hypotenuse, force, distance" };

  // Set the range
  ws["!ref"] = "A1:F3";

  XLSX.utils.book_append_sheet(wb, ws, "Calculations");

  return wb;
}

/**
 * Test File 2: Multi-sheet with cross-sheet references
 *
 * Sheet 1 "Inputs":
 * - A1: length = 5
 * - A2: width = 3
 * - A3: height = 2
 * - B1: density = 7.8 (steel kg/m³ simplified)
 * - B2: price_per_kg = 2.5
 *
 * Sheet 2 "Calculations":
 * - A1: area = Inputs!A1 * Inputs!A2
 * - A2: volume = A1 * Inputs!A3
 * - A3: mass = A2 * Inputs!B1
 * - B1: surface_area = 2*(Inputs!A1*Inputs!A2 + Inputs!A2*Inputs!A3 + Inputs!A1*Inputs!A3)
 * - B2: perimeter = 4*(Inputs!A1 + Inputs!A2 + Inputs!A3)
 *
 * Sheet 3 "Results":
 * - A1: total_cost = Calculations!A3 * Inputs!B2
 * - A2: cost_per_volume = A1 / Calculations!A2
 * - A3: summary = A1 + A2 (just to add another relationship)
 * - B1: area_to_volume_ratio = Calculations!B1 / Calculations!A2
 * - B2: efficiency_score = B1 * 100
 *
 * Isolated Cluster (on Sheet 3):
 * - D1: temp_celsius = 25
 * - D2: temp_fahrenheit = D1 * 9/5 + 32
 * - D3: temp_kelvin = D1 + 273.15
 * - E1: avg_temp = (D1 + D2 + D3) / 3
 * - E2: temp_range = D3 - D1
 */
function createMultiSheetTestFile(): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Inputs
  const wsInputs = XLSX.utils.aoa_to_sheet([
    [5, 7.8], // length, density
    [3, 2.5], // width, price_per_kg
    [2, null], // height
  ]);
  wsInputs["!ref"] = "A1:B3";
  XLSX.utils.book_append_sheet(wb, wsInputs, "Inputs");

  // Sheet 2: Calculations
  const wsCalc = XLSX.utils.aoa_to_sheet([
    [null, null],
    [null, null],
    [null, null],
  ]);

  // Area = length * width
  wsCalc["A1"] = { t: "n", f: "Inputs!A1*Inputs!A2", v: 15 };
  // Volume = area * height
  wsCalc["A2"] = { t: "n", f: "A1*Inputs!A3", v: 30 };
  // Mass = volume * density
  wsCalc["A3"] = { t: "n", f: "A2*Inputs!B1", v: 234 };
  // Surface area = 2(lw + wh + lh)
  wsCalc["B1"] = {
    t: "n",
    f: "2*(Inputs!A1*Inputs!A2+Inputs!A2*Inputs!A3+Inputs!A1*Inputs!A3)",
    v: 62,
  };
  // Perimeter sum = 4(l + w + h)
  wsCalc["B2"] = { t: "n", f: "4*(Inputs!A1+Inputs!A2+Inputs!A3)", v: 40 };

  wsCalc["!ref"] = "A1:B3";
  XLSX.utils.book_append_sheet(wb, wsCalc, "Calculations");

  // Sheet 3: Results
  const wsResults = XLSX.utils.aoa_to_sheet([
    [null, null, null, 25, null],
    [null, null, null, null, null],
    [null, null, null, null, null],
  ]);

  // Main cluster - cost calculations
  // Total cost = mass * price_per_kg
  wsResults["A1"] = { t: "n", f: "Calculations!A3*Inputs!B2", v: 585 };
  // Cost per volume
  wsResults["A2"] = { t: "n", f: "A1/Calculations!A2", v: 19.5 };
  // Summary
  wsResults["A3"] = { t: "n", f: "A1+A2", v: 604.5 };
  // Area to volume ratio
  wsResults["B1"] = { t: "n", f: "Calculations!B1/Calculations!A2", v: 2.067 };
  // Efficiency score
  wsResults["B2"] = { t: "n", f: "B1*100", v: 206.7 };

  // Isolated cluster - temperature conversions
  wsResults["D1"] = { t: "n", v: 25 }; // Celsius (base value)
  wsResults["D2"] = { t: "n", f: "D1*9/5+32", v: 77 }; // Fahrenheit
  wsResults["D3"] = { t: "n", f: "D1+273.15", v: 298.15 }; // Kelvin
  wsResults["E1"] = { t: "n", f: "(D1+D2+D3)/3", v: 133.38 }; // Average
  wsResults["E2"] = { t: "n", f: "D3-D1", v: 248.15 }; // Range

  wsResults["!ref"] = "A1:E3";
  XLSX.utils.book_append_sheet(wb, wsResults, "Results");

  return wb;
}

// Generate and save the files
const simpleWb = createSimpleTestFile();
XLSX.writeFile(simpleWb, path.join(FIXTURES_DIR, "simple-single-sheet.xlsx"));
console.log("Created: simple-single-sheet.xlsx");

const multiWb = createMultiSheetTestFile();
XLSX.writeFile(multiWb, path.join(FIXTURES_DIR, "multi-sheet-cross-ref.xlsx"));
console.log("Created: multi-sheet-cross-ref.xlsx");

console.log("\nTest files generated successfully!");
