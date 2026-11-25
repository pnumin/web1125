"use strict";

const DAY_MS = 86_400_000;
const DAY_RATE = 25_000;
const LODGING_CAP = { seoul: 100_000, metro: 80_000, other: 70_000 };
const numberFormat = new Intl.NumberFormat("ko-KR");
const won = (value) => `${numberFormat.format(value)}원`;

function dayDifference(start, end) {
  return (Date.parse(end) - Date.parse(start)) / DAY_MS;
}

function minutesOfDay(time) {
  const [hour, minute] = time.split(":").map(Number);
  return hour * 60 + minute;
}

function localAllowance(minutes, vehicle) {
  if (vehicle === "assigned") return 0;
  const standard = minutes >= 240 ? 20_000 : 10_000;
  return standard - (vehicle === "reduced" ? 10_000 : 0);
}

function dailyAllowance(days, rate, stayDays) {
  const daysAtTenPercent = Math.max(0, Math.min(stayDays, 31) - 16);
  const daysAtTwentyPercent = Math.max(0, Math.min(stayDays, 61) - 31);
  const daysAtThirtyPercent = Math.max(0, stayDays - 61);
  return days * rate
    - daysAtTenPercent * rate / 10
    - daysAtTwentyPercent * rate / 5
    - daysAtThirtyPercent * rate * 3 / 10;
}

function lodgingAllowance(nights, grade, allowExcess) {
  let standard = 0;
  let extra = 0;
  let unpayable = 0;

  for (const { region, count, cost } of nights) {
    const cap = grade === "second" ? LODGING_CAP[region] : cost;
    const base = Math.min(cost, cap);
    const excess = allowExcess ? Math.min(cost - base, cap * 3 / 10) : 0;
    standard += base * count;
    extra += excess * count;
    unpayable += (cost - base - excess) * count;
  }
  return { standard, extra, unpayable };
}

function setupCalculator() {
  const get = (id) => document.getElementById(id);
  const form = get("trip-form");
  const start = get("start-date");
  const end = get("end-date");
  const grade = get("grade");
  const transport = get("transport");
  const nightList = get("night-list");
  const total = get("total");
  const duration = get("duration");
  const resultDates = get("result-dates");
  const breakdown = get("breakdown");
  const resultNotes = get("result-notes");
  const resultActions = get("result-actions");
  const printButton = get("print-result");
  const downloadButton = get("download-excel");
  const exportStatus = get("export-status");
  let currentResult = null;

  const value = (id) => get(id).value;
  const amount = (id) => get(id).valueAsNumber || 0;
  const checked = (id) => get(id).checked;
  const isLocal = () => form.elements["trip-type"].value === "local";
  const travelDays = () => dayDifference(start.value, end.value) + 1;

  function toggle(id, visible, inputs = []) {
    get(id).hidden = !visible;
    for (const input of inputs) get(input).disabled = !visible;
  }

  function addNight() {
    const node = get("night-template").content.cloneNode(true);
    nightList.append(node);
  }

  function fareInstruction() {
    const transportHints = {
      train: grade.value === "second" ? "일반실 기준의 할인 적용 후 운임을 입력하세요." : "제1호는 특실까지 지급할 수 있습니다. 할인 적용 후 운임을 입력하세요.",
      bus: "정해진 기준과 요율에 따른 버스요금을 입력하세요.",
      flight: "지급 가능한 등급의 운임을 입력하세요. 공적 항공마일리지로 절감된 금액은 운임에서 빼세요.",
      ship: grade.value === "second" ? "2등급 기준의 할인 적용 후 운임을 입력하세요." : "제1호는 1등급까지 지급할 수 있습니다. 할인 적용 후 운임을 입력하세요.",
      car: "해당 경로의 지급 구분에 맞는 철도운임 또는 버스운임을 입력하세요.",
      "car-exception": "기관 기준에서 인정되는 연료비, 통행료 등을 합산해 입력하세요.",
      official: "공용 차량·선박·항공기 이용 구간의 운임은 지급하지 않습니다.",
    };
    return transportHints[transport.value] || "";
  }

  function updateFields() {
    const local = isLocal();
    toggle("local-fields", local, ["local-fields"]);
    toggle("away-fields", !local, ["away-fields"]);
    end.min = start.value || start.min;
    const days = travelDays();
    const validRange = Number.isInteger(days) && days > 0;
    get("meal-days").max = validRange ? days : "";
    get("stay-days").max = validRange ? days : "";

    const hasFare = !["none", "official"].includes(transport.value);
    toggle("fare-fields", transport.value !== "none", ["other-fare"]);
    get("fare").closest("label").hidden = !hasFare;
    get("fare").disabled = !hasFare;
    toggle("other-fare-note", transport.value !== "none");
    const fareHint = fareInstruction();
    get("fare-hint").textContent = fareHint;
    toggle("fare-hint", fareHint !== "");

    const executive = grade.value === "executive";
    toggle("executive-fields", executive, ["actual-daily", "actual-meals"]);
    const reduceLabel = get("reduce-daily").closest("label");
    reduceLabel.hidden = executive;
    get("reduce-daily").disabled = executive;
    for (const id of ["limited-meals", "long-stay"]) {
      get(id).closest("label").hidden = executive;
      get(id).disabled = executive;
    }
    const limitedMeals = !executive && checked("limited-meals");
    toggle("meal-days-label", limitedMeals, ["meal-days"]);
    toggle("meal-days-hint", limitedMeals);
    const longStay = !executive && checked("long-stay");
    toggle("stay-label", longStay, ["stay-days"]);
    toggle("stay-hint", longStay);

    const hasLodging = checked("has-lodging");
    const second = grade.value === "second";
    toggle("lodging-fields", hasLodging);
    toggle("excess-label", hasLodging && second, ["allow-excess"]);
    const rows = [...nightList.children];
    for (const row of rows) {
      const region = row.querySelector(".region");
      region.closest("label").hidden = !second;
      region.disabled = !second || !hasLodging;
      const count = row.querySelector(".night-count");
      count.max = validRange ? days - 1 : "";
      count.disabled = !hasLodging;
      row.querySelector(".night-cost").disabled = !hasLodging;
      row.querySelector(".remove-night").hidden = rows.length === 1;
    }
  }

  function clearResult(message = "일정과 지급 조건을 입력해 주세요.") {
    total.firstChild.nodeValue = "— ";
    duration.textContent = message;
    resultDates.textContent = "";
    resultDates.hidden = true;
    breakdown.replaceChildren();
    resultNotes.replaceChildren();
    breakdown.hidden = true;
    resultNotes.hidden = true;
    resultActions.hidden = true;
    exportStatus.hidden = true;
    currentResult = null;
  }

  function onEdit() {
    get("end-time").setCustomValidity("");
    for (const field of nightList.querySelectorAll(".night-count")) {
      field.setCustomValidity("");
    }
    updateFields();
    if (currentResult) clearResult("조건이 변경되었습니다. 다시 계산해 주세요.");
  }

  function readNights() {
    if (!checked("has-lodging")) return [];
    return [...nightList.children].map((row) => ({
      region: row.querySelector(".region").value,
      count: row.querySelector(".night-count").valueAsNumber,
      cost: row.querySelector(".night-cost").valueAsNumber,
    }));
  }

  function localResult(minutes) {
    const vehicle = value("local-vehicle");
    const notes = [
      minutes >= 240 ? "4시간 이상인 경우 20,000원 (제18조)" : "4시간 미만인 경우 10,000원 (제18조)",
    ];
    if (vehicle === "assigned") notes.push("전용 차량 배정자 지급 제외 (제18조)");
    if (vehicle === "reduced") notes.push("공무용 차량 이용 등으로 10,000원 감액 (제18조)");
    const hours = Math.floor(minutes / 60);
    const remaining = minutes % 60;
    const elapsed = [hours && `${numberFormat.format(hours)}시간`, remaining && `${remaining}분`].filter(Boolean).join(" ");
    return {
      description: `근무지 내 국내 출장 · ${elapsed}`,
      items: [["출장여비", localAllowance(minutes, vehicle)]],
      notes,
    };
  }

  function domesticResult(days, nights) {
    const group = grade.value;
    const executive = group === "executive";
    const reducedRate = !executive && checked("reduce-daily");
    const rate = reducedRate ? DAY_RATE / 2 : DAY_RATE;
    const stayDays = !executive && checked("long-stay") ? amount("stay-days") : 0;
    const mealDays = !executive && checked("limited-meals") ? amount("meal-days") : days;
    const mainFare = ["none", "official"].includes(transport.value) ? 0 : amount("fare");
    const otherFare = transport.value === "none" ? 0 : amount("other-fare");
    const lodging = lodgingAllowance(nights, group, group === "second" && checked("allow-excess"));
    const mileage = amount("mileage");
    const paidNights = nights.reduce((sum, night) => sum + night.count, 0);
    const groupName = { first: "제1호", second: "제2호", executive: "대통령·국무총리" }[group];
    const items = [
      ["운임", mainFare + otherFare],
      ["일비", executive ? amount("actual-daily") : dailyAllowance(days, rate, stayDays)],
      ["숙박비", lodging.standard],
      ["식비", executive ? amount("actual-meals") : mealDays * DAY_RATE],
    ];
    const notes = ["국내 운임·일비·숙박비·식비는 별표 2 적용"];
    if (transport.value === "official") notes.push("공용 교통수단 이용 구간의 운임 제외 (제15조)");
    if (executive) notes.push("일비·식비는 입력한 지급 대상 실비 반영 (별표 2 비고)");
    if (reducedRate) notes.push("일비를 1일 12,500원 기준으로 계산 (제16조 제3항)");
    if (mealDays < days) notes.push(`항공·수로 이동 조건에 따라 식비 ${mealDays}일분 반영 (제16조 제5항)`);
    if (stayDays > 16) notes.push("동일 지역 장기 체재 일비 감액 반영 (제17조)");
    if (lodging.unpayable) notes.push(`숙박 지급 한도를 초과한 ${won(lodging.unpayable)} 제외`);
    if (lodging.extra) {
      items.push(["숙박비 추가 지급", lodging.extra]);
      notes.push("숙박비 추가 지급은 1박 상한액의 30% 이내입니다. 복귀 다음 날부터 1주 이내에 증빙을 갖춰 신청합니다 (제16조 제1·2항).");
    }
    if (mileage) {
      items.push(["공적 항공마일리지 추가 일비", mileage]);
      notes.push("기관 기준에 따라 확인된 추가 일비 입력액 반영 (별표 2 비고)");
    }
    return {
      description: `근무지 외 국내 출장 · ${groupName} · ${numberFormat.format(days)}일 · 지급 대상 숙박 ${numberFormat.format(paidNights)}박`,
      items,
      notes,
    };
  }

  function showResult({ description, items, notes }) {
    const sum = items.reduce((subtotal, [, cost]) => subtotal + cost, 0);
    total.firstChild.nodeValue = `${numberFormat.format(sum)} `;
    duration.textContent = description;
    const local = isLocal();
    const departure = local ? `${start.value} ${value("start-time")}` : start.value;
    const returnDate = local ? `${end.value} ${value("end-time")}` : end.value;
    resultDates.textContent = `출장 일정 ${departure} ~ ${returnDate}`;
    resultDates.hidden = false;
    breakdown.replaceChildren(...items.map(([name, cost]) => {
      const row = document.createElement("div");
      const label = document.createElement("dt");
      const entry = document.createElement("dd");
      label.textContent = name;
      entry.textContent = won(cost);
      row.append(label, entry);
      return row;
    }));
    resultNotes.replaceChildren(...notes.map((text) => {
      const item = document.createElement("li");
      item.textContent = text;
      return item;
    }));
    breakdown.hidden = false;
    resultNotes.hidden = false;
    currentResult = { description, items, notes, sum, departure, returnDate };
    exportStatus.hidden = true;
    resultActions.hidden = false;
    get("result").focus();
  }

  form.addEventListener("input", onEdit);
  printButton.addEventListener("click", () => window.print());
  downloadButton.addEventListener("click", async () => {
    if (!currentResult) return;
    const result = currentResult;
    downloadButton.disabled = true;
    exportStatus.hidden = true;
    try {
      await downloadTripWorkbook(result);
    } catch (error) {
      console.error(error);
      if (result === currentResult) {
        exportStatus.textContent = "엑셀 파일을 저장하지 못했습니다. 다시 시도해 주세요.";
        exportStatus.hidden = false;
      }
    } finally {
      downloadButton.disabled = false;
    }
  });
  form.addEventListener("change", (event) => {
    if (event.target === transport && ["flight", "ship"].includes(transport.value) && grade.value !== "executive") {
      get("limited-meals").checked = true;
      get("special").open = true;
    }
    if (event.target === get("has-lodging") && checked("has-lodging") && !nightList.children.length) {
      addNight();
    }
    onEdit();
  });
  get("add-night").addEventListener("click", () => {
    addNight();
    onEdit();
    nightList.lastElementChild.querySelector("select:not(:disabled), input:not(:disabled)").focus();
  });
  nightList.addEventListener("click", (event) => {
    if (!event.target.matches(".remove-night")) return;
    event.target.closest(".night").remove();
    onEdit();
    get("add-night").focus();
  });
  form.addEventListener("invalid", (event) => {
    const details = event.target.closest("details");
    if (details) details.open = true;
  }, true);
  form.addEventListener("submit", (event) => {
    event.preventDefault();
    if (isLocal()) {
      const minutes = (travelDays() - 1) * 1440
        + minutesOfDay(value("end-time")) - minutesOfDay(value("start-time"));
      if (minutes <= 0) {
        get("end-time").setCustomValidity("복귀 시각은 출발 시각보다 늦어야 합니다.");
        get("end-time").reportValidity();
        return;
      }
      showResult(localResult(minutes));
      return;
    }

    const days = travelDays();
    const nights = readNights();
    if (nights.reduce((sum, night) => sum + night.count, 0) > days - 1) {
      const count = nightList.lastElementChild.querySelector(".night-count");
      count.setCustomValidity(`전체 숙박 박수는 입력한 일정 기준 ${days - 1}박을 넘을 수 없습니다.`);
      count.reportValidity();
      return;
    }
    showResult(domesticResult(days, nights));
  });
  form.addEventListener("reset", () => {
    queueMicrotask(() => {
      nightList.replaceChildren();
      get("special").open = false;
      updateFields();
      clearResult();
    });
  });
  updateFields();
}

setupCalculator();

const rulesDialog = document.getElementById("rules-dialog");
document.getElementById("open-rules").addEventListener("click", () => rulesDialog.showModal());
