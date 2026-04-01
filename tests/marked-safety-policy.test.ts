import assert from "node:assert/strict";
import { test } from "node:test";

import {
  createMarkdownImageRenderPlan,
  getMarkdownImageLabel,
  isAllowedMarkdownUrl,
  isMarkdownExtensionDisabledByPolicy,
} from "../src/compat/marked-safety-policy.ts";

void test("isAllowedMarkdownUrl blocks unsafe protocols", () => {
  assert.equal(isAllowedMarkdownUrl("https://example.com"), true);
  assert.equal(isAllowedMarkdownUrl("http://example.com"), true);
  assert.equal(isAllowedMarkdownUrl("mailto:test@example.com"), true);
  assert.equal(isAllowedMarkdownUrl("tel:+12025550123"), true);
  assert.equal(isAllowedMarkdownUrl("#cite:Sheet1!A1:B10"), true);

  assert.equal(isAllowedMarkdownUrl("javascript:alert(1)"), false);
  assert.equal(isAllowedMarkdownUrl("data:text/html,<h1>x</h1>"), false);
  assert.equal(isAllowedMarkdownUrl("file:///etc/passwd"), false);
  assert.equal(isAllowedMarkdownUrl(""), false);
});

void test("policy disables all KaTeX math extensions", () => {
  assert.equal(isMarkdownExtensionDisabledByPolicy("inlineMathDollar"), true);
  assert.equal(isMarkdownExtensionDisabledByPolicy("blockMathDollar"), true);
  assert.equal(isMarkdownExtensionDisabledByPolicy("inlineMathLatex"), true);
  assert.equal(isMarkdownExtensionDisabledByPolicy("blockMathLatex"), true);
  assert.equal(isMarkdownExtensionDisabledByPolicy("someOtherExtension"), false);
});

void test("markdown image plans never render <img>", () => {
  const safePlan = createMarkdownImageRenderPlan("https://cdn.example.com/logo.png", "Logo");
  assert.deepEqual(safePlan, {
    kind: "link",
    href: "https://cdn.example.com/logo.png",
    label: "image: Logo",
  });

  const unsafePlan = createMarkdownImageRenderPlan("javascript:alert(1)", "Pwn");
  assert.deepEqual(unsafePlan, {
    kind: "text",
    label: "image: Pwn",
  });

  assert.equal(getMarkdownImageLabel(""), "image");
});
