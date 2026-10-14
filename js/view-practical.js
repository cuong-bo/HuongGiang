// Practical (essay) exams: a list of the seven sets and a detail page per task
// with the transcribed statement and a step-by-step guide.
(function () {
  const { h, icon } = window.Dom;
  const Core = window.Core;

  const APP_ICONS = { word: "fileText", excel: "table", powerpoint: "presentation" };
  const COPIED_RESET_MS = 1600;
  const MISSING_NOTICE =
    "Đề bài của câu này chưa có vì ảnh chụp bị thiếu. Phần hướng dẫn bên dưới chỉ là bước chung, sẽ được viết thêm khi có đề.";

  const sets = () => window.PRACTICAL_DATA || [];
  const isMissing = (task) => task.statement.requirements.length === 0;
  const taskHref = (setId, no) => `#/thuc-hanh/${setId}/${no}`;
  const timeText = (set) => (set.timeLimitMinutes ? `${set.timeLimitMinutes} phút` : "Không ghi thời gian");

  /* ---------- list page ---------- */

  function taskRow(set, task) {
    return h(
      "a",
      { class: "task-row", href: taskHref(set.id, task.no), "data-testid": `task-link-${set.id}-${task.no}` },
      h("span", { class: "app-badge" }, icon(APP_ICONS[task.app])),
      h(
        "span",
        { class: "task-row-body" },
        h("strong", {}, `Câu ${task.no}, ${Core.appLabel(task.app)}`),
        h("span", { class: "task-row-intro" }, isMissing(task) ? "Chưa có đề bài (thiếu ảnh)" : task.statement.intro)
      ),
      icon("arrowRight")
    );
  }

  function setCard(set) {
    return h(
      "article",
      { class: "de-card", "data-testid": `de-card-${set.id}` },
      h("div", { class: "de-head" }, h("h2", {}, `Đề số ${set.id}`), h("span", { class: "muted" }, timeText(set))),
      h("div", { class: "de-tasks" }, set.tasks.map((task) => taskRow(set, task)))
    );
  }

  function renderList() {
    const all = sets();
    const taskCount = all.reduce((n, s) => n + s.tasks.length, 0);
    return h(
      "div",
      { class: "practical wrap" },
      h("h1", { class: "page-title" }, "Hướng dẫn thi thực hành"),
      h(
        "p",
        { class: "lede" },
        `${all.length} đề, ${taskCount} câu Word, Excel, PowerPoint. Mỗi câu có đề bài dạng chữ và hướng dẫn từng bước.`
      ),
      h(
        "p",
        { class: "muted source-note" },
        "Đề được chuyển từ ảnh chụp và tài liệu ôn thi. Chỗ ảnh bị mờ hoặc thiếu được ghi chú ở cuối mỗi câu."
      ),
      h("div", { class: "de-grid", "data-testid": "de-grid" }, all.map(setCard))
    );
  }

  /* ---------- detail page ---------- */

  async function copyText(text) {
    try {
      await navigator.clipboard.writeText(text);
      return true;
    } catch (err) {
      const area = document.createElement("textarea");
      area.value = text;
      area.setAttribute("readonly", "");
      area.className = "copy-fallback";
      document.body.append(area);
      area.select();
      let ok = false;
      try {
        ok = document.execCommand("copy");
      } catch (e) {
        ok = false;
      }
      area.remove();
      return ok;
    }
  }

  function materialBlock(block) {
    if (block.kind === "table") {
      const [head, ...body] = block.rows;
      return h(
        "figure",
        { class: "material" },
        h("figcaption", {}, block.caption),
        h(
          "div",
          { class: "table-scroll", role: "region", tabindex: "0", "aria-label": block.caption },
          h(
            "table",
            {},
            h("thead", {}, h("tr", {}, head.map((cell) => h("th", { scope: "col" }, cell)))),
            h("tbody", {}, body.map((row) => h("tr", {}, row.map((cell) => h("td", {}, cell)))))
          )
        )
      );
    }
    return h(
      "figure",
      { class: "material" },
      h("figcaption", {}, block.caption),
      h("div", { class: "doc-lines" }, block.lines.map((line) => h("p", {}, line || " ")))
    );
  }

  function statementSection(task) {
    const s = task.statement;
    return h(
      "section",
      { class: "panel", id: "sec-statement", tabindex: "-1" },
      h("h2", {}, "Đề bài"),
      isMissing(task) ? h("p", { class: "notice", "data-testid": "missing-notice" }, MISSING_NOTICE) : null,
      s.intro ? h("p", { class: "statement-intro" }, s.intro) : null,
      s.requirements.length > 0
        ? h("ul", { class: "req-list" }, s.requirements.map((r) => h("li", {}, r)))
        : null,
      s.material.map(materialBlock)
    );
  }

  function stepCard(step, index) {
    return h(
      "li",
      { class: "step", "data-testid": "step" },
      h(
        "div",
        { class: "step-head" },
        h("span", { class: "step-num", "aria-hidden": "true" }, index + 1),
        h("h3", {}, step.title),
        step.forRequirement ? h("span", { class: "step-for" }, `Cho yêu cầu: ${step.forRequirement}`) : null
      ),
      h("ol", { class: "actions" }, step.actions.map((a) => h("li", {}, a))),
      step.note ? h("p", { class: "note" }, h("strong", {}, "Lưu ý: "), step.note) : null
    );
  }

  function solutionSection(task) {
    const sol = task.solution;
    return h(
      "section",
      { class: "panel", id: "sec-solution", tabindex: "-1" },
      h("h2", {}, "Hướng dẫn giải"),
      h("p", { class: "statement-intro" }, sol.summary),
      h("ol", { class: "steps" }, sol.steps.map(stepCard))
    );
  }

  function formulaCard(app, f) {
    const status = h("span", { class: "sr-only", role: "status", "aria-live": "polite" });
    const label = h("span", {}, "Chép công thức");
    const button = h(
      "button",
      { type: "button", class: "copy-btn", "data-testid": "copy-formula" },
      icon("copy"),
      label,
      status
    );
    button.addEventListener("click", async () => {
      const ok = await copyText(f.formula);
      label.textContent = ok ? "Đã chép" : "Không chép được, hãy bôi đen để chép";
      status.textContent = ok ? "Đã chép công thức" : "Chép công thức thất bại";
      const id = setTimeout(() => {
        label.textContent = "Chép công thức";
        status.textContent = "";
      }, COPIED_RESET_MS);
      app.onCleanup(() => clearTimeout(id));
    });
    return h(
      "li",
      { class: "formula", "data-testid": "formula" },
      h("div", { class: "formula-head" }, h("strong", {}, `Ô ${f.cell}`), button),
      h("code", { class: "formula-code" }, f.formula),
      h("p", { class: "muted" }, f.meaning)
    );
  }

  function formulasSection(app, task) {
    const list = task.solution.formulas;
    if (list.length === 0) return null;
    return h(
      "section",
      { class: "panel", id: "sec-formulas", tabindex: "-1" },
      h("h2", {}, "Công thức Excel"),
      h(
        "p",
        { class: "muted" },
        "Nếu máy dùng dấu chấm phẩy để ngăn cách đối số hàm, hãy đổi dấu phẩy thành dấu chấm phẩy khi nhập."
      ),
      h("ul", { class: "formulas" }, list.map((f) => formulaCard(app, f)))
    );
  }

  function listSection(id, title, items, className) {
    if (items.length === 0) return null;
    return h("section", { class: "panel", id, tabindex: "-1" }, h("h2", {}, title), h("ul", { class: className }, items));
  }

  function checklistSection(task) {
    const items = task.checklist.map((text) =>
      h("li", {}, h("label", {}, h("input", { type: "checkbox" }), h("span", {}, text)))
    );
    return listSection("sec-check", "Tự kiểm tra", items, "check-list");
  }

  function mistakesSection(task) {
    const items = task.commonMistakes.map((text) => h("li", {}, text));
    return listSection("sec-mistakes", "Lỗi thường gặp", items, "bullet-list");
  }

  function resultSection(task) {
    if (!task.solution.result) return null;
    return h(
      "section",
      { class: "panel", id: "sec-result", tabindex: "-1" },
      h("h2", {}, "Kết quả mong đợi"),
      h("p", {}, task.solution.result)
    );
  }

  function uncertainSection(task) {
    if (task.uncertain.length === 0) return null;
    return h(
      "details",
      { class: "uncertain", "data-testid": "uncertain" },
      h("summary", {}, icon("warning"), "Ghi chú về độ chính xác của đề", icon("caretDown")),
      h("ul", { class: "bullet-list" }, task.uncertain.map((text) => h("li", {}, text)))
    );
  }

  function outline(task) {
    const entries = [
      ["sec-statement", "Đề bài"],
      ["sec-solution", "Hướng dẫn giải"],
      task.solution.formulas.length > 0 ? ["sec-formulas", "Công thức Excel"] : null,
      task.solution.result ? ["sec-result", "Kết quả mong đợi"] : null,
      ["sec-check", "Tự kiểm tra"],
      task.commonMistakes.length > 0 ? ["sec-mistakes", "Lỗi thường gặp"] : null,
    ].filter(Boolean);
    return h(
      "nav",
      { class: "outline", "aria-label": "Trong trang này" },
      h("p", { class: "outline-title" }, "Trong trang này"),
      entries.map(([id, text]) =>
        h(
          "button",
          {
            type: "button",
            class: "outline-link",
            onclick: () => {
              const el = document.getElementById(id);
              if (el) {
                el.scrollIntoView();
                el.focus({ preventScroll: true });
              }
            },
          },
          text
        )
      )
    );
  }

  function neighborLink(label, target, dir) {
    if (!target) return h("span", {});
    const next = Core.findTask(sets(), target.setId, target.no);
    return h(
      "a",
      { class: `btn btn-ghost nav-${dir}`, href: taskHref(target.setId, target.no), "data-testid": `${dir}-task` },
      dir === "prev" ? icon("arrowLeft") : null,
      h("span", { class: "nav-text" }, h("small", {}, label), `Đề ${next.set.id}, Câu ${next.task.no}`),
      dir === "next" ? icon("arrowRight") : null
    );
  }

  function renderTask(app) {
    const { setId, no } = app.route;
    const found = Core.findTask(sets(), setId, no);
    const { set, task } = found;
    const near = Core.taskNeighbors(sets(), setId, no);
    return h(
      "div",
      { class: "task-page wrap" },
      h(
        "a",
        { class: "back-link", href: "#/thuc-hanh", "data-testid": "back-to-list" },
        icon("arrowLeft"),
        "Danh sách đề"
      ),
      h(
        "header",
        { class: "task-head" },
        h("p", { class: "muted" }, `${set.title}, ${timeText(set)}`),
        h("h1", { class: "page-title" }, task.title)
      ),
      set.generalNotes.length > 0
        ? h(
            "details",
            { class: "general-notes" },
            h("summary", {}, "Quy định chung của đề", icon("caretDown")),
            h("ul", { class: "bullet-list" }, set.generalNotes.map((n) => h("li", {}, n)))
          )
        : null,
      h(
        "div",
        { class: "task-grid" },
        h(
          "div",
          { class: "task-main" },
          statementSection(task),
          solutionSection(task),
          formulasSection(app, task),
          resultSection(task),
          checklistSection(task),
          mistakesSection(task),
          uncertainSection(task)
        ),
        outline(task)
      ),
      h(
        "div",
        { class: "task-nav" },
        neighborLink("Câu trước", near.prev, "prev"),
        neighborLink("Câu tiếp", near.next, "next")
      )
    );
  }

  window.Views = window.Views || {};
  window.Views.practical = renderList;
  window.Views.task = renderTask;
})();
