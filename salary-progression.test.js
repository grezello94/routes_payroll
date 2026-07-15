const assert = require("node:assert/strict");
const { effectiveSalary } = require("./salary-progression");

assert.equal(effectiveSalary(15000, [
  { month: "2026-06", presentSalary: 15000, increment: 1000 },
], "2026-06"), 16000, "the award month includes its increment");

assert.equal(effectiveSalary(15000, [
  { month: "2026-06", presentSalary: 15000, increment: 1000 },
], "2026-07", false), 16000, "the next month carries the raise forward");

assert.equal(effectiveSalary(15000, [
  { month: "2026-06", presentSalary: 15000, increment: 1000 },
  { month: "2026-07", presentSalary: 15000, increment: 0 },
], "2026-08", false), 16000, "a stale future row cannot erase a raise");

assert.equal(effectiveSalary(15000, [
  { month: "2026-06", presentSalary: 15000, increment: 1000 },
  { month: "2026-07", presentSalary: 16000, increment: 500 },
], "2026-08", false), 16500, "separate later raises accumulate through checkpoints");

assert.equal(effectiveSalary(15000, [
  { month: "2026-12", presentSalary: 15000, increment: 1000 },
], "2027-01", false), 16000, "raises carry across year boundaries");

assert.equal(effectiveSalary(15000, [
  { month: "2026-06", presentSalary: 15000, increment: 1000 },
  { month: "2026-07", presentSalary: 15000, increment: 1000 },
], "2026-07"), 16000, "a repeated stale checkpoint is not double-counted");

console.log("Salary progression regression tests passed.");
