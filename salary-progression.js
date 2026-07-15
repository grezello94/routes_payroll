(function exposeSalaryProgression(root, factory) {
  const api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  if (root) root.SalaryProgression = api;
}(typeof globalThis !== "undefined" ? globalThis : this, function createSalaryProgression() {
  function nonNegativeMoney(value) {
    const amount = Number(value || 0);
    return Number.isFinite(amount) ? Math.max(0, amount) : 0;
  }

  function effectiveSalary(baseSalary, rows, cutoffMonth, includeCutoff = true) {
    let result = nonNegativeMoney(baseSalary);
    for (const row of Array.isArray(rows) ? rows : []) {
      const rowMonth = String(row?.month || "");
      if (!rowMonth) continue;
      if (includeCutoff ? rowMonth > cutoffMonth : rowMonth >= cutoffMonth) continue;
      const presentSalary = nonNegativeMoney(
        row.presentSalary ?? row.present_salary ?? row.baseSalary ?? row.base_salary
      );
      const checkpoint = presentSalary + nonNegativeMoney(row.increment);
      result = Math.max(result, checkpoint);
    }
    return result;
  }

  return { effectiveSalary };
}));
