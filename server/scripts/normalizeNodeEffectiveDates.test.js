import { test } from "node:test";
import assert from "node:assert/strict";
import { planNodeNormalization } from "./normalizeNodeEffectiveDates.js";

test("flattens a node with varying effectiveFrom onto its earliest date", () => {
  const rows = [
    { clientId: "test", nodeId: "entity:a", nodeHistoryKey: "NODE#entity:a#AT#2025-12-31#VER#v1", versionId: "v1", effectiveFrom: "2025-12-31", effectiveTo: "2026-02-01", changedAt: "2025-12-31T00:00:00.000Z", changedBy: "import" },
    { clientId: "test", nodeId: "entity:a", nodeHistoryKey: "NODE#entity:a#AT#2026-02-01#VER#v2", versionId: "v2", effectiveFrom: "2026-02-01", effectiveTo: "2026-05-15", changedAt: "2026-02-01T10:00:00.000Z", changedBy: "alice" },
    { clientId: "test", nodeId: "entity:a", nodeHistoryKey: "NODE#entity:a#AT#2026-05-15#VER#v3", versionId: "v3", effectiveFrom: "2026-05-15", changedAt: "2026-05-15T12:00:00.000Z", changedBy: "bob" },
  ];

  const plan = planNodeNormalization(rows, { nodeId: "entity:a" });
  assert.ok(plan);
  assert.equal(plan.originDate, "2025-12-31");
  assert.equal(plan.writes.length, 3); // all three rows need a key/effectiveFrom change

  const byVersionId = new Map(plan.writes.map((w) => [w.item.versionId, w]));

  const v1 = byVersionId.get("v1");
  assert.equal(v1.oldKey, null); // effectiveFrom unchanged, so key is unchanged
  assert.equal(v1.item.effectiveFrom, "2025-12-31");
  assert.equal(v1.item.effectiveTo, "2025-12-31"); // closed, zero-width
  assert.equal(v1.item.changedBy, "import"); // audit fields untouched

  const v2 = byVersionId.get("v2");
  assert.equal(v2.oldKey, "NODE#entity:a#AT#2026-02-01#VER#v2");
  assert.equal(v2.item.nodeHistoryKey, "NODE#entity:a#AT#2025-12-31#VER#v2");
  assert.equal(v2.item.effectiveFrom, "2025-12-31");
  assert.equal(v2.item.effectiveTo, "2025-12-31");
  assert.equal(v2.item.changedBy, "alice");

  const v3 = byVersionId.get("v3");
  assert.equal(v3.oldKey, "NODE#entity:a#AT#2026-05-15#VER#v3");
  assert.equal(v3.item.effectiveFrom, "2025-12-31");
  assert.equal(v3.item.effectiveTo, undefined); // the most recent edit stays open
  assert.equal(v3.item.changedBy, "bob");
});

test("is a no-op for a node whose history already shares one effectiveFrom", () => {
  const rows = [
    { clientId: "test", nodeId: "entity:b", nodeHistoryKey: "NODE#entity:b#AT#2024-01-01#VER#v1", versionId: "v1", effectiveFrom: "2024-01-01", effectiveTo: "2024-01-01", changedAt: "2024-01-01T00:00:00.000Z" },
    { clientId: "test", nodeId: "entity:b", nodeHistoryKey: "NODE#entity:b#AT#2024-01-01#VER#v2", versionId: "v2", effectiveFrom: "2024-01-01", changedAt: "2025-03-01T00:00:00.000Z" },
  ];
  assert.equal(planNodeNormalization(rows, { nodeId: "entity:b" }), null);
});

test("is a no-op for a node with only a single history row", () => {
  const rows = [
    { clientId: "test", nodeId: "entity:c", nodeHistoryKey: "NODE#entity:c#AT#2024-06-01#VER#v1", versionId: "v1", effectiveFrom: "2024-06-01", changedAt: "2024-06-01T00:00:00.000Z" },
  ];
  assert.equal(planNodeNormalization(rows, { nodeId: "entity:c" }), null);
});

test("returns null when no row has a usable effectiveFrom", () => {
  const rows = [
    { clientId: "test", nodeId: "entity:d", nodeHistoryKey: "NODE#entity:d#AT#undefined#VER#v1", versionId: "v1", changedAt: "2024-06-01T00:00:00.000Z" },
  ];
  assert.equal(planNodeNormalization(rows, { nodeId: "entity:d" }), null);
});

test("correctly orders by changedAt even when array order is scrambled", () => {
  const rows = [
    { clientId: "test", nodeId: "entity:e", nodeHistoryKey: "k3", versionId: "v3", effectiveFrom: "2026-05-15", changedAt: "2026-05-15T12:00:00.000Z" },
    { clientId: "test", nodeId: "entity:e", nodeHistoryKey: "k1", versionId: "v1", effectiveFrom: "2025-12-31", changedAt: "2025-12-31T00:00:00.000Z" },
    { clientId: "test", nodeId: "entity:e", nodeHistoryKey: "k2", versionId: "v2", effectiveFrom: "2026-02-01", changedAt: "2026-02-01T10:00:00.000Z" },
  ];
  const plan = planNodeNormalization(rows, { nodeId: "entity:e" });
  const openRow = plan.writes.find((w) => !w.item.effectiveTo);
  assert.equal(openRow.item.versionId, "v3"); // most recent changedAt, regardless of input order
});
