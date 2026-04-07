const SHEET_ID = "1Z_pS4LOoVCb4A2eevPp9NczvhCUWpKkLq3zWJEImdSM"; 
  const GID = "0";           

  // Polling refresh: checks for updates every N ms
  const REFRESH_EVERY_MS = 30000; // 30s

// Post height to parent whenever content changes
function sendHeight() {
  const height = document.body.scrollHeight;
  window.parent.postMessage({ iframeHeight: height }, '*');
}

// Send on load and whenever content might change
window.addEventListener('load', sendHeight);
setInterval(sendHeight, 500);

function gvizUrl(){
  const tqx = "out:json;headers=1";
  return `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?gid=${encodeURIComponent(GID)}&tqx=${encodeURIComponent(tqx)}&tq=${encodeURIComponent("select *")}`;
}

// Known header label values to skip if they sneak through as a data row
const HEADER_VALUES = new Set([
  "first name", "firstname", "surname", "last name",
  "jobtitle", "job title", "organisation", "organization",
  "category", "linkedin", "imageurl", "image url",
  "a", "b", "c", "d", "e", "f", "g"
]);

function isHeaderRow(r) {
  // If the "FirstName" or "Surname" cell looks like a column header label, skip it
  const fn = (r.FirstName ?? "").toString().trim().toLowerCase();
  const sn = (r.Surname ?? "").toString().trim().toLowerCase();
  return HEADER_VALUES.has(fn) || HEADER_VALUES.has(sn) ||
         fn === "first name" || fn === "firstname" ||
         sn === "surname" || sn === "last name";
}

async function fetchSheetRows(){
  const res = await fetch(gvizUrl(), { cache: "no-store" });
  if(!res.ok) throw new Error(`Fetch failed (${res.status})`);
  const text = await res.text();

  const json = JSON.parse(text.substring(text.indexOf("{"), text.lastIndexOf("}") + 1));
  const table = json.table;

  // Use label if present, otherwise fall back to column id (A, B, C...)
  const headers = table.cols.map(c => (c.label || c.id || "").trim());

  const rows = table.rows.map(r => {
    const obj = {};
    headers.forEach((h, i) => {
      const cell = r.c[i];
      obj[h] = cell ? (cell.v ?? "") : "";
    });
    return obj;
  });

  const hasRealHeaders = headers.some(h => h === "First Name" || h === "FirstName" || h === "Surname");

  let mapped;
  if (!hasRealHeaders) {
    mapped = rows.map(r => ({
      FirstName: r.A ?? "",
      Surname: r.B ?? "",
      JobTitle: r.C ?? "",
      Organisation: r.D ?? "",
      Category: r.E ?? "",
      LinkedIn: r.F ?? "",
      ImageURL: r.G ?? "",
    }));
  } else {
    mapped = rows.map(r => ({
      FirstName: r["First Name"] ?? r["FirstName"] ?? r.A ?? "",
      Surname: r["Surname"] ?? r["Last Name"] ?? r.B ?? "",
      JobTitle: r["JobTitle"] ?? r["Job Title"] ?? r.C ?? "",
      Organisation: r["Organisation"] ?? r["Organization"] ?? r.D ?? "",
      Category: r["Category"] ?? r.E ?? "",
      LinkedIn: r["LinkedIn"] ?? r.F ?? "",
      ImageURL: r["ImageURL"] ?? r["Image URL"] ?? r.G ?? "",
    }));
  }

  // Drop any row that looks like the header row itself
  return mapped.filter(r => !isHeaderRow(r));
}


  /******************************************************************
   * STATE + FILTERING
   ******************************************************************/
  let allPeople = [];
  let currentLetter = "ALL";
  let currentQuery = "";
  let lastHash = "";

  function normalize(s){
    return (s ?? "").toString().trim();
  }

  function getInitials(firstName, surname){
    const a = firstName?.[0] ?? "";
    const b = surname?.[0] ?? (firstName?.[1] ?? "");
    return (a + (b || "")).toUpperCase() || "?";
  }

  function applyFilters(){
    const q = currentQuery.toLowerCase();
    const filtered = allPeople.filter(p => {
      const fullName = (normalize(p.FirstName) + " " + normalize(p.Surname)).trim();
      const matchesQuery = !q || fullName.toLowerCase().includes(q);

      const letter = currentLetter;
      const matchesLetter =
        (letter === "ALL") ||
        (normalize(p.Surname)?.[0]?.toUpperCase() === letter) ||
        (!normalize(p.Surname) && normalize(p.FirstName)?.[0]?.toUpperCase() === letter);

      return matchesQuery && matchesLetter;
    });

    render(filtered);
  }


  const gridEl = document.getElementById("grid");
  const emptyEl = document.getElementById("empty");

  function chainIcon(){
    return `
      <svg class="chain" viewBox="0 0 24 24" fill="none" aria-hidden="true">
        <path d="M10 13a5 5 0 0 1 0-7l1.2-1.2a5 5 0 0 1 7 7L17 13"
              stroke="currentColor" stroke-width="2" stroke-linecap="round"/>
        <path d="M14 11a5 5 0 0 1 0 7L12.8 19.2a5 5 0 0 1-7-7L7 11"
              stroke="currentColor" stroke-width="2" stroke-linecap="round"/>
      </svg>
    `;
  }

  function render(list){
    gridEl.innerHTML = "";

    if(list.length === 0){
      emptyEl.style.display = "block";
    } else {
      emptyEl.style.display = "none";
    }

    // Sort A→Z by surname, then first name
    const sorted = [...list].sort((a,b) => {
      const surnameA = normalize(a.Surname).toLowerCase();
      const surnameB = normalize(b.Surname).toLowerCase();
      if(surnameA !== surnameB) return surnameA.localeCompare(surnameB);
      return normalize(a.FirstName).toLowerCase().localeCompare(normalize(b.FirstName).toLowerCase());
    });

    for(const p of sorted){
      const firstName = normalize(p.FirstName);
      const surname = normalize(p.Surname);
      const fullName = [firstName, surname].filter(Boolean).join(" ");

      const job = normalize(p.JobTitle);
      const org = normalize(p.Organisation);
      const category = normalize(p.Category);
      const linkedinRaw = normalize(p.LinkedIn);
      const img = normalize(p.ImageURL);

      // Ensure LinkedIn URL is absolute
      let safeLinkedIn = "";
      if(linkedinRaw){
        if(linkedinRaw.startsWith("http://") || linkedinRaw.startsWith("https://")){
          safeLinkedIn = linkedinRaw;
        } else {
          safeLinkedIn = "https://" + linkedinRaw;
        }
      }

      const card = document.createElement("div");
      card.className = "card";

      const photo = img
        ? `<img src="${escapeHtml(img)}" alt="${escapeHtml(fullName)}" loading="lazy">`
        : `<div class="placeholder" aria-label="${escapeHtml(fullName)}">${escapeHtml(getInitials(firstName, surname))}</div>`;

      card.innerHTML = `
        <div class="photoWrap">
          <div class="photoRing">
            ${photo}
          </div>
        </div>

        <div class="name">
          <span class="first">${escapeHtml(firstName)}</span>
          <span class="last">${escapeHtml(surname)}</span>
        </div>

        <div class="rows">
          <div class="row"><em>${escapeHtml(job || "Job title")}</em></div>
          <div class="row">${escapeHtml(org || "Organisation")}</div>
          ${category ? `<div class="row small">${escapeHtml(category)}</div>` : ""}
        </div>

        <div class="links">
          ${chainIcon()}
          ${safeLinkedIn
            ? `<a href="${escapeHtml(safeLinkedIn)}" target="_blank" rel="noopener noreferrer">LinkedIn</a>`
            : `<span style="color:rgba(11,107,46,.65);font-family:var(--serif);">LinkedIn</span>`
          }
        </div>
      `;

      // If image fails to load, swap to initials placeholder
      const imgEl = card.querySelector("img");
      if(imgEl){
        imgEl.addEventListener("error", () => {
          const ring = card.querySelector(".photoRing");
          ring.innerHTML = `<div class="placeholder">${escapeHtml(getInitials(firstName, surname))}</div>`;
        }, { once: true });
      }

      gridEl.appendChild(card);
    }
  }

  function escapeHtml(str){
    return (str ?? "").toString()
      .replaceAll("&","&amp;")
      .replaceAll("<","&lt;")
      .replaceAll(">","&gt;")
      .replaceAll('"',"&quot;")
      .replaceAll("'","&#039;");
  }

  /******************************************************************
   * A–Z BAR
   ******************************************************************/
  const azBarEl = document.getElementById("azBar");

  function buildAz(){
    const letters = ["ALL", ..."ABCDEFGHIJKLMNOPQRSTUVWXYZ".split("")];
    azBarEl.innerHTML = "";

    for(const L of letters){
      const btn = document.createElement("button");
      btn.type = "button";
      btn.textContent = (L === "ALL") ? "•" : L;
      btn.title = (L === "ALL") ? "Show all" : `Show ${L}`;
      btn.className = (currentLetter === L) ? "active" : "";
      btn.addEventListener("click", () => {
        currentLetter = L;
        [...azBarEl.querySelectorAll("button")].forEach(b => b.classList.remove("active"));
        btn.classList.add("active");
        applyFilters();
      });
      azBarEl.appendChild(btn);
    }
  }

  /******************************************************************
   * SEARCH
   ******************************************************************/
  const searchInputEl = document.getElementById("searchInput");
  let searchTimer = null;

  searchInputEl.addEventListener("input", () => {
    clearTimeout(searchTimer);
    searchTimer = setTimeout(() => {
      currentQuery = searchInputEl.value || "";
      applyFilters();
    }, 120);
  });

  /******************************************************************
   * AUTO-REFRESH (poll + re-render only if changed)
   ******************************************************************/
  function simpleHash(s){
    let h = 0;
    for(let i=0; i<s.length; i++){
      h = (h << 5) - h + s.charCodeAt(i);
      h |= 0;
    }
    return String(h);
  }

  async function loadAndMaybeUpdate({force=false} = {}){
    try{
      const rows = await fetchSheetRows();

      const cleaned = rows
        .map(r => ({
          FirstName: normalize(r.FirstName),
          Surname: normalize(r.Surname),
          JobTitle: normalize(r.JobTitle),
          Organisation: normalize(r.Organisation),
          Category: normalize(r.Category),
          LinkedIn: normalize(r.LinkedIn),
          ImageURL: normalize(r.ImageURL),
        }))
        .filter(r => r.FirstName || r.Surname || r.JobTitle || r.Organisation);

      const hash = simpleHash(JSON.stringify(cleaned));

      if(force || hash !== lastHash){
        lastHash = hash;
        allPeople = cleaned;
        applyFilters();
      }
    } catch (err){
      console.error(err);
    }
  }


  buildAz();
  loadAndMaybeUpdate({force:true});
  setInterval(() => loadAndMaybeUpdate({force:false}), REFRESH_EVERY_MS);


  buildAz();
  loadAndMaybeUpdate({force:true});
  setInterval(() => loadAndMaybeUpdate({force:false}), REFRESH_EVERY_MS);
