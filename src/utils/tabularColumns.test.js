import assert from "node:assert/strict";
import test from "node:test";
import { normalizeSavedColumnOrder, resolveTabularColumns } from "./tabularColumns.js";

for (const baseKey of ["name", "owner"]) {
  test(`${baseKey}: saved dictionary columns survive loading and saving`, () => {
    const savedOrder = [baseKey, "dd:new-field", "dd:deleted-field"];
    const normalizedOrder = normalizeSavedColumnOrder(savedOrder);
    const baseColumns = [{ key: baseKey }];
    assert.deepEqual(normalizedOrder, savedOrder);
    assert.deepEqual(resolveTabularColumns(normalizedOrder, baseColumns), baseColumns);

    const dictionaryColumn = { key: "dd:new-field", label: "New field" };
    const loadedColumns = [...baseColumns, dictionaryColumn];
    assert.deepEqual(
      resolveTabularColumns(normalizedOrder, loadedColumns),
      [baseColumns[0], dictionaryColumn]
    );
    assert.deepEqual(normalizeSavedColumnOrder(normalizedOrder), savedOrder);
    assert.deepEqual(savedOrder, [baseKey, "dd:new-field", "dd:deleted-field"]);
  });
}

test("legacy normalization and hidden columns work before dictionary loading", () => {
  assert.deepEqual(
    normalizeSavedColumnOrder(
      [" name ", "dd:dd:new-field", "dd:dd:hidden-field", "address"],
      ["dd:hidden-field", "address"]
    ),
    ["name", "dd:new-field"]
  );
});

test("invalid orders and keys do not introduce columns", () => {
  assert.deepEqual(normalizeSavedColumnOrder(null), []);
  assert.deepEqual(normalizeSavedColumnOrder({ name: true }), []);
  assert.deepEqual(
    normalizeSavedColumnOrder(["", " ", null, 4, {}, "name"], null),
    ["name"]
  );
});

test("rendering preserves order, removes duplicates, and skips unavailable definitions", () => {
  const columns = [{ key: "name" }, { key: "dd:new-field" }, { key: "address" }];
  assert.deepEqual(
    resolveTabularColumns(
      ["dd:new-field", "missing", "address", "dd:new-field", "name"],
      columns
    ),
    [columns[1], columns[2], columns[0]]
  );
  assert.deepEqual(resolveTabularColumns(undefined, columns), []);
});
