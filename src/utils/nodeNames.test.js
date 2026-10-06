import test from "node:test";
import assert from "node:assert/strict";
import { describeNameConflict, findNameConflicts, normalizeNodeName } from "./nodeNames.js";

const index = [
  { id: "c|entity:acme", kind: "entity", currentName: "Acme, Inc.", deleted: false, names: [
    { name: "Acme, Inc.", kind: "entity", effectiveFrom: "2024-01-01", effectiveTo: null },
  ] },
  { id: "c|entity:beta", kind: "entity", currentName: "Beta Partners", deleted: false, names: [
    { name: "Gamma LLC", kind: "entity", effectiveFrom: "2023-01-01", effectiveTo: "2025-04-01" },
    { name: "Beta Partners", kind: "entity", effectiveFrom: "2025-04-01", effectiveTo: null },
  ] },
  { id: "c|entity:old", kind: "entity", currentName: null, deleted: true, names: [
    { name: "Delta Co", kind: "entity", effectiveFrom: "2022-01-01", effectiveTo: null },
  ] },
  { id: "c|person:acme", kind: "person", currentName: "Acme Inc", deleted: false, names: [] },
];

test("normalizeNodeName ignores case, spacing and punctuation", () => {
  assert.equal(normalizeNodeName("  Acme,  Inc. "), "acme inc");
  assert.equal(normalizeNodeName("Café Ünïon"), "café ünïon");
  assert.equal(normalizeNodeName("..."), "");
});

test("findNameConflicts matches current names of the same kind only", () => {
  const conflicts = findNameConflicts(index, { name: "acme inc", kind: "entity" });
  assert.deepEqual(conflicts.map((c) => [c.id, c.type]), [["c|entity:acme", "current"]]);
  assert.match(describeNameConflict(conflicts[0]), /An existing entity is named "Acme, Inc."/);
});

test("findNameConflicts reports former names and deleted nodes", () => {
  const past = findNameConflicts(index, { name: "Gamma LLC", kind: "entity" });
  assert.equal(past[0].type, "past");
  assert.equal(past[0].to, "2025-04-01");
  assert.match(describeNameConflict(past[0]), /former name of the entity now named "Beta Partners" \(from 01\/01\/2023 through 03\/31\/2025\)/);

  const deleted = findNameConflicts(index, { name: "delta co", kind: "entity" });
  assert.equal(deleted[0].type, "deleted");
  assert.match(describeNameConflict(deleted[0]), /A deleted entity was named "Delta Co"/);
});

test("findNameConflicts excludes the node being edited and uses extra nodes", () => {
  assert.equal(findNameConflicts(index, { name: "Gamma LLC", kind: "entity", excludeId: "c|entity:beta" }).length, 0);
  const extra = findNameConflicts([], { name: "New Co", kind: "entity", extraNodes: [{ id: "x", kind: "entity", name: "New Co" }] });
  assert.equal(extra[0].type, "current");
  assert.equal(findNameConflicts(index, { name: "", kind: "entity" }).length, 0);
});
