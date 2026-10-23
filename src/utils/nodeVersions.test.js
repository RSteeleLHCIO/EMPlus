import test from "node:test";
import assert from "node:assert/strict";
import {
  addDaysIso,
  describeSaveRange,
  findVersionIndexAt,
  formatVersionRange,
  getChangedVersionFields,
  getVersionStatus,
} from "./nodeVersions.js";

const versions = [
  { effectiveFrom: "2025-01-01", effectiveTo: "2025-06-01", name: "A" },
  { effectiveFrom: "2025-06-01", effectiveTo: "2026-01-01", name: "B" },
  { effectiveFrom: "2026-01-01", effectiveTo: null, name: "C" },
];

test("findVersionIndexAt picks the version in effect (end date exclusive)", () => {
  assert.equal(findVersionIndexAt(versions, "2025-05-31"), 0);
  assert.equal(findVersionIndexAt(versions, "2025-06-01"), 1);
  assert.equal(findVersionIndexAt(versions, "2030-01-01"), 2);
  assert.equal(findVersionIndexAt(versions, "2024-01-01"), 0);
  assert.equal(findVersionIndexAt([], "2025-01-01"), -1);
});

test("findVersionIndexAt falls back to the prior version inside a gap", () => {
  const gapped = [
    { effectiveFrom: "2025-01-01", effectiveTo: "2025-03-01" },
    { effectiveFrom: "2025-06-01", effectiveTo: null },
  ];
  assert.equal(findVersionIndexAt(gapped, "2025-04-15"), 0);
});

test("getVersionStatus classifies current, historical and future", () => {
  assert.equal(getVersionStatus(versions[2], "2026-10-06"), "current");
  assert.equal(getVersionStatus(versions[0], "2026-10-06"), "historical");
  assert.equal(getVersionStatus({ effectiveFrom: "2027-01-01" }, "2026-10-06"), "future");
});

test("formatVersionRange shows an inclusive end date", () => {
  assert.equal(formatVersionRange("2025-01-01", "2025-06-01"), "01/01/2025 – 05/31/2025");
  assert.equal(formatVersionRange("2026-01-01", null), "01/01/2026 – present");
  assert.equal(addDaysIso("2025-03-01", -1), "2025-02-28");
});

test("getChangedVersionFields reports changed built-in and custom fields", () => {
  const changed = getChangedVersionFields(
    { name: "A", address: "", emails: ["a@x.com"], customFields: { f1: "x", f2: ["1"] } },
    { name: "A", address: "1 Main", emails: "a@x.com", customFields: { f1: "y", f2: ["1"], f3: "" } }
  );
  assert.deepEqual([...changed.keys()].sort(), ["address", "cf:f1"]);
  assert.equal(changed.get("cf:f1"), "x");
});

test("getChangedVersionFields tracks notes and treats a missing attribute as blank", () => {
  assert.equal(getChangedVersionFields({ name: "A" }, { name: "A", notes: "" }).size, 0);
  const changed = getChangedVersionFields({ name: "A" }, { name: "A", notes: "Call Dr. K" });
  assert.deepEqual([...changed.keys()], ["notes"]);
  assert.equal(changed.get("notes"), undefined);
});

test("describeSaveRange mirrors server version placement", () => {
  const correction = describeSaveRange(versions, "2025-06-01");
  assert.equal(correction.replaces, versions[1]);
  assert.equal(correction.to, "2026-01-01");
  assert.equal(correction.isLatest, false);

  const split = describeSaveRange(versions, "2025-09-01");
  assert.equal(split.replaces, null);
  assert.equal(split.splits, versions[1]);
  assert.equal(split.to, "2026-01-01");

  const latest = describeSaveRange(versions, "2026-10-06");
  assert.equal(latest.isLatest, true);
  assert.equal(latest.splits, versions[2]);
});
