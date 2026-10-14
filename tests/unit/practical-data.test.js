// Guards the transcribed practical exams: shape, completeness and house style.
const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const DIR = path.join(__dirname, "..", "..", "data", "practical");
const SETS = fs
  .readdirSync(DIR)
  .filter((f) => /^de\d+\.json$/.test(f))
  .map((f) => JSON.parse(fs.readFileSync(path.join(DIR, f), "utf8")))
  .sort((a, b) => a.id - b.id);

const APPS = new Set(["word", "excel", "powerpoint"]);
const DASHES = /[–—]/;

function allStrings(value, out = []) {
  if (typeof value === "string") out.push(value);
  else if (Array.isArray(value)) value.forEach((v) => allStrings(v, out));
  else if (value && typeof value === "object") Object.values(value).forEach((v) => allStrings(v, out));
  return out;
}

test("there are seven sets numbered 1 to 7", () => {
  assert.deepEqual(
    SETS.map((s) => s.id),
    [1, 2, 3, 4, 5, 6, 7]
  );
});

for (const set of SETS) {
  test(`set ${set.id}: has a title and three tasks with valid apps`, () => {
    assert.ok(set.title.length > 0);
    assert.equal(set.tasks.length, 3);
    for (const task of set.tasks) assert.ok(APPS.has(task.app), `bad app ${task.app}`);
    assert.deepEqual(
      [...set.tasks.map((t) => t.app)].sort(),
      ["excel", "powerpoint", "word"]
    );
  });

  test(`set ${set.id}: every task has statement, steps and checklist`, () => {
    for (const task of set.tasks) {
      assert.ok(task.title.length > 0);
      assert.ok(task.statement && typeof task.statement.intro === "string");
      assert.ok(Array.isArray(task.statement.requirements));
      assert.ok(Array.isArray(task.statement.material));
      assert.ok(task.solution.steps.length > 0, `task ${task.no} has no steps`);
      for (const step of task.solution.steps) {
        assert.ok(step.title.length > 0);
        assert.ok(step.actions.length > 0, `step "${step.title}" has no actions`);
      }
      assert.ok(task.checklist.length > 0);
      assert.ok(Array.isArray(task.commonMistakes));
      assert.ok(Array.isArray(task.uncertain));
    }
  });

  test(`set ${set.id}: formulas only on excel tasks, and every excel task has some`, () => {
    for (const task of set.tasks) {
      const formulas = task.solution.formulas;
      if (task.app !== "excel") {
        assert.equal(formulas.length, 0, `task ${task.no} is not excel`);
        continue;
      }
      assert.ok(formulas.length > 0, `excel task ${task.no} has no formulas`);
      for (const f of formulas) {
        assert.ok(f.cell && f.formula.startsWith("="), `bad formula ${JSON.stringify(f)}`);
        assert.ok(f.meaning.length > 0);
      }
    }
  });

  test(`set ${set.id}: material blocks are well formed`, () => {
    for (const task of set.tasks) {
      for (const block of task.statement.material) {
        assert.ok(["text", "table"].includes(block.kind));
        if (block.kind === "text") assert.ok(block.lines.length > 0);
        else assert.ok(block.rows.length > 0 && block.rows.every((r) => Array.isArray(r)));
      }
    }
  });

  test(`set ${set.id}: no em or en dashes anywhere`, () => {
    const offenders = allStrings(set).filter((s) => DASHES.test(s));
    assert.deepEqual(offenders, []);
  });
}
