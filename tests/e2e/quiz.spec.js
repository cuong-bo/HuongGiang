const { test, expect } = require("@playwright/test");

// Module 1, Q1 "Phần cứng máy tính là gì?" has its correct answer at index 2 (C).
const CORRECT = 2;
const WRONG = 0;

test.beforeEach(async ({ page }) => {
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("console", (m) => m.type() === "error" && errors.push(m.text()));
  page.errors = errors;
  await page.goto("/");
  await page.evaluate(() => {
    localStorage.clear();
    sessionStorage.clear();
  });
  await page.reload();
});

test.afterEach(async ({ page }) => {
  expect(page.errors, "no console or page errors").toEqual([]);
});

test("home lists the six modules with the total question count", async ({ page }) => {
  await expect(page).toHaveTitle(/Quý ông người pháp/);
  await expect(page.getByRole("link", { name: /Quý ông người pháp/ })).toBeVisible();
  await expect(page.getByTestId("module-grid").locator("article")).toHaveCount(6);
  await expect(page.getByRole("heading", { level: 1 })).toContainText("Ôn thi CNTT cơ bản");
  await expect(page.locator(".lede")).toContainText("718 câu hỏi");
  await expect(page.getByTestId("wrong-empty")).toBeVisible();
});

test("a wrong answer shows the correct answer and locks the question", async ({ page }) => {
  await page.getByTestId("module-card-1").getByRole("link").click();
  await expect(page.getByTestId("q-text")).toContainText("Phần cứng máy tính là gì?");

  await page.getByTestId(`option-${WRONG}`).click();

  await expect(page.getByTestId("feedback")).toContainText("Chưa đúng");
  await expect(page.getByTestId("feedback")).toContainText("Đáp án đúng là C");
  await expect(page.getByTestId(`option-${WRONG}`)).toHaveAttribute("data-state", "wrong");
  await expect(page.getByTestId(`option-${CORRECT}`)).toHaveAttribute("data-state", "correct");
  await expect(page.getByTestId("tally-bad")).toHaveText("1");

  // Locked: choosing another option does not change the result.
  // The option is aria-disabled, so Playwright's actionability check would refuse; force the click.
  await page.getByTestId("option-1").click({ force: true });
  await expect(page.getByTestId("option-1")).toHaveAttribute("data-state", "idle");
  await expect(page.getByTestId("tally-bad")).toHaveText("1");
  await expect(page.getByTestId("pal-1")).toHaveClass(/is-bad/);
});

test("a correct answer is confirmed and progress survives a reload", async ({ page }) => {
  await page.goto("/#/practice/1");
  await page.getByTestId(`option-${CORRECT}`).click();
  await expect(page.getByTestId("feedback")).toContainText("Chính xác");
  await expect(page.getByTestId("tally-ok")).toHaveText("1");

  await page.reload();
  // Resumes at the first unanswered question (Q2) and keeps the tally.
  await expect(page.getByTestId("q-position")).toContainText("Câu 2 trên 120");
  await expect(page.getByTestId("tally-ok")).toHaveText("1");
  await expect(page.getByTestId("pal-1")).toHaveClass(/is-ok/);
});

test("keyboard: letter keys answer and arrow keys move between questions", async ({ page }) => {
  await page.goto("/#/practice/1");
  await page.keyboard.press("c");
  await expect(page.getByTestId(`option-${CORRECT}`)).toHaveAttribute("data-state", "correct");
  await page.keyboard.press("ArrowRight");
  await expect(page.getByTestId("q-position")).toContainText("Câu 2 trên 120");
  await page.keyboard.press("ArrowLeft");
  await expect(page.getByTestId("q-position")).toContainText("Câu 1 trên 120");
});

test("the question palette jumps to a question", async ({ page }) => {
  await page.setViewportSize({ width: 1280, height: 900 });
  await page.goto("/#/practice/2");
  await page.getByTestId("pal-10").click();
  await expect(page.getByTestId("q-position")).toContainText("Câu 10 trên 120");
});

test("wrong answers land in the review list and leave it once answered correctly", async ({ page }) => {
  await page.goto("/#/practice/1");
  await page.getByTestId(`option-${WRONG}`).click();
  await expect(page.getByTestId("wrong-count")).toHaveText("1");

  await page.getByRole("link", { name: /Câu sai/ }).click();
  await expect(page.getByRole("heading", { level: 1 })).toHaveText("Ôn lại câu sai");
  await expect(page.getByTestId("q-position")).toContainText("Câu 1 trên 1");

  await page.getByTestId(`option-${CORRECT}`).click();
  await expect(page.getByTestId("feedback")).toContainText("Chính xác");
  await expect(page.getByTestId("wrong-count")).toHaveCount(0);
});

test("exam hides feedback until submit, then reviews wrong answers", async ({ page }) => {
  await page.getByTestId("start-exam").click();
  await expect(page).toHaveURL(/#\/exam$/);
  await expect(page.getByTestId("timer")).toHaveText(/^\d\d:\d\d$/);

  await page.getByTestId("option-0").click();
  await expect(page.getByTestId("option-0")).toHaveAttribute("data-state", "selected");
  await expect(page.getByTestId("feedback")).toBeEmpty();

  await page.getByTestId("submit-exam").click();
  await expect(page.getByRole("dialog")).toContainText("chưa trả lời");
  await page.getByTestId("dialog-confirm").click();

  await expect(page).toHaveURL(/#\/result$/);
  await expect(page.getByRole("heading", { level: 1 })).toHaveText("Kết quả thi thử");
  await expect(page.getByTestId("score")).toHaveText(/^\d+\/30$/);
  expect(await page.getByTestId("review-item").count()).toBeGreaterThan(0);

  await page.getByTestId("filter-all").check();
  await expect(page.getByTestId("review-item")).toHaveCount(30);

  await page.getByTestId("retry-wrong").click();
  await expect(page.getByRole("heading", { level: 1 })).toHaveText("Ôn lại câu sai");
  // Unanswered questions are retried too: 30 total, minus at most the one answered correctly.
  await expect(page.getByTestId("q-position")).toContainText(/Câu 1 trên (29|30)$/);
});

test("a first load on an in-page anchor still renders the home page", async ({ page }) => {
  await page.goto("/#main");
  await expect(page.getByTestId("module-grid")).toBeVisible();
});

test("the server survives a malformed URL", async ({ request }) => {
  const bad = await request.get("/%");
  expect(bad.status()).toBe(400);
  const ok = await request.get("/");
  expect(ok.status()).toBe(200);
});

test("progress made in another tab appears here without a reload", async ({ page, context }) => {
  const other = await context.newPage();
  await other.goto("/#/practice/1");
  await other.getByTestId(`option-${CORRECT}`).click();
  await expect(page.getByTestId("module-card-1")).toContainText("Đã làm 1/120");
  await other.close();
});

test("an exam in progress can be resumed after leaving", async ({ page }) => {
  await page.getByTestId("start-exam").click();
  await page.getByTestId("option-1").click();
  await page.getByTestId("exit-quiz").click();
  await page.getByTestId("resume-exam").getByRole("link").click();
  await expect(page.getByTestId("option-1")).toHaveAttribute("data-state", "selected");
});

test("resetting a module asks for confirmation and clears its progress", async ({ page }) => {
  await page.goto("/#/practice/1");
  await page.getByTestId(`option-${CORRECT}`).click();
  await page.getByTestId("exit-quiz").click();

  await page.getByTestId("module-reset-1").click();
  await page.getByRole("dialog").getByRole("button", { name: "Hủy" }).click();
  await expect(page.getByTestId("module-card-1")).toContainText("Đã làm 1/120");

  await page.getByTestId("module-reset-1").click();
  await page.getByTestId("dialog-confirm").click();
  await expect(page.getByTestId("module-card-1")).toContainText("Chưa bắt đầu");
});

test("theme toggle switches and remembers the choice", async ({ page }) => {
  const before = await page.locator("html").getAttribute("data-theme");
  await page.getByTestId("theme-toggle").click();
  const after = await page.locator("html").getAttribute("data-theme");
  expect(after).not.toBe(before);
  await page.reload();
  await expect(page.locator("html")).toHaveAttribute("data-theme", after);
});

test("the last question of a module offers a finish action", async ({ page }) => {
  await page.goto("/#/practice/5");
  await page.getByTestId("pal-118").click();
  await expect(page.getByTestId("next")).toHaveText("Hoàn thành");
  await expect(page.getByTestId("option-1")).toBeVisible();
  await expect(page.getByTestId("option-2")).toHaveCount(0);
});

test.describe("practical exams", () => {
  test("the home page links to the list of seven sets and 21 tasks", async ({ page }) => {
    await page.getByTestId("cta-practical").click();
    await expect(page).toHaveURL(/#\/thuc-hanh$/);
    await expect(page.getByRole("heading", { level: 1 })).toHaveText("Hướng dẫn thi thực hành");
    await expect(page.getByTestId("de-grid").locator("article")).toHaveCount(7);
    await expect(page.locator("[data-testid^=task-link-]")).toHaveCount(21);
    await expect(page.getByRole("link", { name: "Thực hành" })).toHaveAttribute("aria-current", "page");
  });

  test("a task shows the transcribed statement, ordered steps and Excel formulas", async ({ page }) => {
    await page.goto("/#/thuc-hanh");
    await page.getByTestId("task-link-5-2").click();
    await expect(page.getByRole("heading", { level: 1 })).toContainText("Câu 2. Tạo bảng tính");
    await expect(page.getByText("Quá trình học tập môn Lập trình cơ bản")).toBeVisible();
    expect(await page.getByTestId("step").count()).toBeGreaterThanOrEqual(8);
    await expect(page.getByTestId("formula").first()).toContainText("=ROUND(AVERAGE(D3:F3),1)");
    await expect(page.getByRole("heading", { name: "Tự kiểm tra" })).toBeVisible();
    await expect(page).toHaveTitle(/Đề 5, Câu 2/);
  });

  test("a step with real screenshots shows them with captions and an enlarge link", async ({ page }) => {
    await page.goto("/#/thuc-hanh/1/1");
    const shots = page.getByTestId("shot");
    await expect(shots).toHaveCount(30);
    const first = shots.first();
    await expect(first.locator("img")).toHaveAttribute("src", /assets\/img\/practical\/de1-c1\/b1-1\.png$/);
    await expect(first.locator("a")).toHaveAttribute("target", "_blank");
    await expect(first.locator("figcaption")).not.toBeEmpty();
    await first.locator("img").scrollIntoViewIfNeeded();
    const loaded = await first.locator("img").evaluate((img) => img.complete && img.naturalWidth > 0);
    expect(loaded).toBe(true);
  });

  test("screenshots do not make the page scroll sideways on a phone", async ({ page }) => {
    await page.setViewportSize({ width: 375, height: 800 });
    await page.goto("/#/thuc-hanh/1/1");
    await expect(page.getByTestId("shot").first()).toBeVisible();
    const overflow = await page.evaluate(() => document.documentElement.scrollWidth > window.innerWidth);
    expect(overflow).toBe(false);
  });

  test("copying a formula confirms to the user", async ({ page, context }) => {
    await context.grantPermissions(["clipboard-read", "clipboard-write"]);
    await page.goto("/#/thuc-hanh/5/2");
    const first = page.getByTestId("formula").first();
    await first.getByTestId("copy-formula").click();
    await expect(first.getByTestId("copy-formula")).toContainText("Đã chép");
  });

  test("previous and next walk across tasks and the back link returns to the list", async ({ page }) => {
    await page.goto("/#/thuc-hanh/5/2");
    await page.getByTestId("next-task").click();
    await expect(page).toHaveURL(/#\/thuc-hanh\/5\/3$/);
    await page.getByTestId("prev-task").click();
    await expect(page).toHaveURL(/#\/thuc-hanh\/5\/2$/);
    await page.getByTestId("back-to-list").click();
    await expect(page).toHaveURL(/#\/thuc-hanh$/);
  });

  test("a task whose statement image is missing says so honestly", async ({ page }) => {
    await page.goto("/#/thuc-hanh/4/3");
    await expect(page.getByTestId("missing-notice")).toContainText("chưa có");
  });

  test("tasks with unreadable image parts list them in a notes section", async ({ page }) => {
    await page.goto("/#/thuc-hanh/1/2");
    const notes = page.getByTestId("uncertain");
    await expect(notes).toBeVisible();
    await notes.locator("summary").click();
    await expect(notes.locator("li").first()).toBeVisible();
  });

  test("an unknown task falls back to the list", async ({ page }) => {
    await page.goto("/#/thuc-hanh/99/1");
    await expect(page.getByTestId("de-grid")).toBeVisible();
  });

  test("the detail page does not scroll sideways on a phone", async ({ page }) => {
    await page.setViewportSize({ width: 375, height: 800 });
    await page.goto("/#/thuc-hanh/5/2");
    await expect(page.getByTestId("formula").first()).toBeVisible();
    const overflow = await page.evaluate(() => document.documentElement.scrollWidth > window.innerWidth);
    expect(overflow).toBe(false);
  });
});
