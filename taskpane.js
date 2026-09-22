
"use strict";

const $ = id => document.getElementById(id);

let wordReady = false;
let busy = false;

const colors = {
  yellow: "#FFFF00",
  green: "#00FF00",
  cyan: "#00FFFF",
  magenta: "#FF00FF",
  red: "#FF0000",
  blue: "#0000FF",
  gray: "#C0C0C0"
};

const formatIds = [
  "bold", "italic", "underline",
  "font", "size", "changeFontColor",
  "fontColor", "highlight"
];

function status(message) {
  $("status").textContent = message;
}

function getTerms() {
  const terms = $("terms").value
    .split("\n")
    .map(x => x.trim())
    .filter(Boolean);

  return [...new Set(terms)];
}

function settings() {
  const result = {};

  for (const id of formatIds) {
    const element = $(id);

    result[id] = element.type === "checkbox"
      ? element.checked
      : element.value;
  }

  return result;
}

function applySettings(s) {
  for (const id of formatIds) {
    if (!(id in s)) continue;

    const element = $(id);

    if (element.type === "checkbox") {
      element.checked = Boolean(s[id]);
    } else {
      element.value = s[id];
    }
  }
}

function getPresets() {
  try {
    return JSON.parse(
      localStorage.getItem("findFormatPresets") || "{}"
    );
  } catch {
    return {};
  }
}

function refreshPresets() {
  const select = $("presets");

  select.replaceChildren(
    new Option("Choose a preset", "")
  );

  for (const name of Object.keys(getPresets()).sort()) {
    select.add(new Option(name, name));
  }
}

function savePreset() {
  const name = $("presetName").value.trim();

  if (!name) {
    status("Enter a name for the preset.");
    return;
  }

  const presets = getPresets();

  presets[name] = settings();

  try {
    localStorage.setItem(
      "findFormatPresets",
      JSON.stringify(presets)
    );

    refreshPresets();
    $("presets").value = name;

    status("Preset saved: " + name);
  } catch (error) {
    status("Could not save preset: " + error.message);
  }
}

function loadPreset() {
  const name = $("presets").value;
  const preset = getPresets()[name];

  if (!preset) {
    status("Choose a saved preset first.");
    return;
  }

  applySettings(preset);
  $("presetName").value = name;

  status("Preset loaded: " + name);
}

function deletePreset() {
  const name = $("presets").value;

  if (!name) {
    status("Choose a preset to delete.");
    return;
  }

  const presets = getPresets();

  delete presets[name];

  localStorage.setItem(
    "findFormatPresets",
    JSON.stringify(presets)
  );

  refreshPresets();

  status("Preset deleted: " + name);
}

function formatRange(range, s) {
  const font = range.font;

  if (s.bold !== "keep") {
    font.bold = s.bold === "on";
  }

  if (s.italic !== "keep") {
    font.italic = s.italic === "on";
  }

  if (s.underline !== "keep") {
    font.underline = s.underline === "on"
      ? "Single"
      : "None";
  }

  if (s.font) {
    font.name = s.font;
  }

  if (s.size !== "") {
    const size = Number(s.size);

    if (!Number.isFinite(size) ||
        size < 1 || size > 1638) {
      throw new Error("Invalid font size.");
    }

    font.size = size;
  }

  if (s.changeFontColor) {
    font.color = s.fontColor;
  }

  if (s.highlight !== "keep") {
    font.highlightColor = s.highlight === "none"
      ? null
      : colors[s.highlight];
  }
}

function validateSearch() {
  if (!wordReady) {
    throw new Error(
      "Open this add-in inside Microsoft Word."
    );
  }

  const terms = getTerms();

  if (!terms.length) {
    throw new Error(
      "Enter at least one word or phrase."
    );
  }

  if ($("regex").checked &&
      $("scope").value === "selection") {
    throw new Error(
      "Regular expressions currently require " +
      "the entire document scope."
    );
  }

  if ($("regex").checked) {
    const flags = $("matchCase").checked ? "g" : "gi";

    for (const term of terms) {
      new RegExp(term, flags);
    }
  }

  return terms;
}

async function ordinarySearch(
  context, terms, shouldApply, s
) {
  const source = $("scope").value === "selection"
    ? context.document.getSelection()
    : context.document.body;

  const searches = [];

  for (const term of terms) {
    const results = source.search(term, {
      matchCase: $("matchCase").checked,
      matchWholeWord: $("wholeWord").checked,
      matchWildcards: false
    });

    results.load("items");

    searches.push(results);
  }

  await context.sync();

  let count = 0;

  for (const results of searches) {
    for (const range of results.items) {
      count++;

      if (shouldApply) {
        formatRange(range, s);
      }
    }
  }

  if (shouldApply) {
    await context.sync();
  }

  return count;
}

/*
 * Regular-expression search:
 *
 * Searches individual body paragraphs.
 * Regex matches must not span paragraph boundaries.
 * Zero-length matches are ignored.
 *
 * Each matching string is located using Word's
 * native search API before formatting is applied.
 */

async function regexSearch(
  context, terms, shouldApply, s
) {
  const paragraphs = context.document.body.paragraphs;

  paragraphs.load("items/text");

  await context.sync();

  const flags = $("matchCase").checked ? "g" : "gi";

  const patterns = terms.map(
    term => new RegExp(term, flags)
  );

  let count = 0;

  for (const paragraph of paragraphs.items) {
    const text = paragraph.text;

    if (!text) continue;

    const matches = [];

    for (const pattern of patterns) {
      pattern.lastIndex = 0;

      for (const match of text.matchAll(pattern)) {
        if (!match[0].length) continue;

        matches.push({
          start: match.index,
          text: match[0]
        });
      }
    }

    if (!matches.length) continue;

    /*
     * Remove duplicate matches generated by
     * overlapping search patterns.
     */

    const unique = new Map();

    for (const match of matches) {
      unique.set(
        match.start + ":" + match.text.length,
        match
      );
    }

    const filtered = [...unique.values()];

    /*
     * Group matches by exact matching text.
     */

    const groups = new Map();

    for (const match of filtered) {
      if (!groups.has(match.text)) {
        groups.set(match.text, []);
      }

      groups.get(match.text).push(match.start);
    }

    for (const [literal, positions] of groups) {
      const results = paragraph.search(literal, {
        matchCase: true,
        matchWholeWord: false,
        matchWildcards: false
      });

      results.load("items");

      await context.sync();

      /*
       * Word search results are returned in
       * document order. Match their occurrences
       * to character positions in paragraph text.
       */

      const occurrences = [];

      let position = 0;

      while (position < text.length) {
        const index = text.indexOf(literal, position);

        if (index < 0) break;

        occurrences.push(index);

        position = index + literal.length;
      }

      const selected = new Set(positions);

      for (
        let i = 0;
        i < results.items.length &&
        i < occurrences.length;
        i++
      ) {
        if (!selected.has(occurrences[i])) {
          continue;
        }

        count++;

        if (shouldApply) {
          formatRange(results.items[i], s);
        }
      }
    }
  }

  if (shouldApply) {
    await context.sync();
  }

  return count;
}

async function runSearch(shouldApply) {
  if (busy) return;

  busy = true;

  $("findButton").disabled = true;
  $("applyButton").disabled = true;

  try {
    const terms = validateSearch();
    const s = settings();

    status(
      shouldApply
        ? "Applying formatting..."
        : "Searching document..."
    );

    const count = await Word.run(async context => {
      if ($("regex").checked) {
        return await regexSearch(
          context, terms, shouldApply, s
        );
      }

      return await ordinarySearch(
        context, terms, shouldApply, s
      );
    });

    if (shouldApply) {
      status(
        "Formatting complete.\n" +
        count + " matches processed."
      );
    } else {
      status(
        "Search complete.\n" +
        count + " matches found."
      );
    }

  } catch (error) {
    console.error(error);

    status(
      "Error: " +
      (error.message || String(error))
    );

  } finally {
    busy = false;

    $("findButton").disabled = false;
    $("applyButton").disabled = false;
  }
}

$("savePreset").addEventListener(
  "click", savePreset
);

$("loadPreset").addEventListener(
  "click", loadPreset
);

$("deletePreset").addEventListener(
  "click", deletePreset
);

$("findButton").addEventListener(
  "click", () => runSearch(false)
);

$("applyButton").addEventListener(
  "click", () => runSearch(true)
);

refreshPresets();

Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    wordReady = true;

    status("Ready. Enter your search terms.");
  } else {
    status(
      "Open Find & Format inside Microsoft Word."
    );
  }
});
