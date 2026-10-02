(() => {
    "use strict";

    // Configuration and shared service definitions.
    const CONFIG = {
        excelFile: "CNX-LoS/CNX LoS.xlsx",
        excelUrl: "https://o365cmu-my.sharepoint.com/:x:/g/personal/salintip_ain_cmu_ac_th/IQBKHTSvvaTARISYnNkCFvu9AW7ueXj1gUnASjGdjB3BCSs?e=yT0DPf",
        dashboardYears: [2025, 2026],
        monthNamesEn: [
            "", "January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December"
        ]
    };
    const monthNames = {
        1: "มกราคม",
        2: "กุมภาพันธ์",
        3: "มีนาคม",
        4: "เมษายน",
        5: "พฤษภาคม",
        6: "มิถุนายน",
        7: "กรกฎาคม",
        8: "สิงหาคม",
        9: "กันยายน",
        10: "ตุลาคม",
        11: "พฤศจิกายน",
        12: "ธันวาคม"
    };
    const SERVICES = {
        Departure: ["Check-In", "In-Line Screening", "Security Screening"],
        Arrival: ["Immigration", "Baggage Claim"],
        "Arrival-International": ["Immigration", "Baggage Claim"],
        "Arrival-Domestic": ["Baggage Claim", "Taxi"]
    };
    const ALL_SERVICES = [...SERVICES.Departure, ...SERVICES.Arrival];
    const FILTER_IDS = ["year", "month", "date", "flight", "direction", "service"];
    const numberFormatter = new Intl.NumberFormat("en-US");
    const state = { rows: [] };
    const elements = {};

    // Value conversion and presentation. Durations are stored in minutes.
    function toNumeric(value) {
        if (value === null || value === undefined || value === "") return 0;
        if (typeof value === "number") return Number.isFinite(value) ? value : 0;
        const cleaned = String(value).replace(/,/g, "").replace(/[^0-9.-]/g, "");
        const number = Number(cleaned);
        return Number.isFinite(number) ? number : 0;
    }

    function safeText(value) {
        return value === null || value === undefined ? "" : String(value).trim();
    }

    function escapeHtml(value) {
        const entities = { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" };
        return String(value ?? "").replace(/[&<>"']/g, character => entities[character]);
    }

    function parseTimeMinutes(value) {
        const match = String(value || "").match(/^(\d{1,2}):(\d{2})$/);
        return match ? Number(match[1]) * 60 + Number(match[2]) : null;
    }

    function calculateWaitingMinutes(startValue, finishValue, fallbackValue = 0) {
        const start = parseTimeMinutes(startValue);
        const finish = parseTimeMinutes(finishValue);
        if (start === null || finish === null) return toNumeric(fallbackValue);
        return finish >= start ? finish - start : finish + 1440 - start;
    }

    function formatWaitingDuration(value) {
        if (value === null || value === undefined || value === "") return "-";
        const minutes = Number(value);
        if (!Number.isFinite(minutes) || minutes < 0) return "-";
        const totalSeconds = Math.round(minutes * 60);
        const m = Math.floor(totalSeconds / 60);
        const s = totalSeconds % 60;
        return `${m}:${String(s).padStart(2, "0")}`;
    }

    function formatNumber(value) {
        return numberFormatter.format(Number(value || 0));
    }

    function formatDateThai(dateString) {
        if (!dateString) return "-";
        const date = new Date(dateString);
        if (Number.isNaN(date.getTime())) return dateString;
        return `${date.getDate()}/${date.getMonth() + 1}/${date.getFullYear()}`;
    }

    // Excel parsing: preserve the workbook's existing column and date conventions.
    function normalizeRow(row) {
        if (!row || typeof row !== "object") return null;

        const columns = Object.keys(row);
        const readColumn = aliases => {
            for (const alias of aliases) {
                const column = columns.find(name => name.toLowerCase() === alias.toLowerCase());
                if (column) return row[column];
            }
            return null;
        };
        const yearRaw = readColumn(["year", "ปี"]);
        const monthRaw = readColumn(["month", "เดือน"]);
        const dateRaw = readColumn(["date", "วันที่", "day"]);
        const flightRaw = readColumn(["flight", "flight type", "type", "ประเภท"]);
        const directionRaw = readColumn(["direction", "dir", "ขาเข้า/ขาออก", "Arrival/Departure"]);
        const serviceRaw = readColumn(["service", "service area", "service area name", "Area"]);
        const startRaw = readColumn(["start", "start time", "เวลาเริ่ม"]);
        const finishRaw = readColumn(["finish", "end time", "finish time", "เวลาสิ้นสุด"]);
        const waitingRaw = readColumn(["waiting", "waiting time", "avg waiting time", "waiting time (min)"]);
        const passengerRaw = readColumn(["passenger", "passengers", "passenger count"]);

        const yearMatch = String(yearRaw || "").match(/\d{4}/);
        const year = Number(yearMatch?.[0] || new Date().getFullYear());
        let month = parseMonthValue(monthRaw);
        if (!month && dateRaw) {
            const date = new Date(String(dateRaw));
            if (!Number.isNaN(date.getTime())) month = date.getMonth() + 1;
        }
        const service = safeText(serviceRaw || "Unknown");
        if (!service) return null;
        const start = normalizeTime(startRaw);
        const finish = normalizeTime(finishRaw);

        return {
            year,
            month: month || 1,
            monthLabel: monthNames[month] || "เดือน",
            date: parseDateValue(dateRaw || "", year, month || 1),
            flight: normalizeFlight(flightRaw || "International"),
            direction: normalizeDirection(directionRaw || "Arrival"),
            service,
            start,
            finish,
            waiting: calculateWaitingMinutes(start, finish, waitingRaw),
            passenger: toNumeric(passengerRaw)
        };
    }

    function parseMonthValue(value) {
        if (
            value === null ||
            value === undefined ||
            value === ""
        ) {
            return null;
        }
        if (typeof value === "number") {
            return (
                Number.isInteger(value) &&
                value >= 1 &&
                value <= 12
            )
                ? value
                : null;
        }
        const s = String(value)
            .trim()
            .toLowerCase();
        const direct = Number(s);
        if (
            Number.isInteger(direct) &&
            direct >= 1 &&
            direct <= 12
        ) {
            return direct;
        }
        const monthMap = {
            jan: 1,
            january: 1,
            feb: 2,
            february: 2,
            mar: 3,
            march: 3,
            apr: 4,
            april: 4,
            may: 5,
            jun: 6,
            june: 6,
            jul: 7,
            july: 7,
            aug: 8,
            august: 8,
            sep: 9,
            sept: 9,
            september: 9,
            oct: 10,
            october: 10,
            nov: 11,
            november: 11,
            dec: 12,
            december: 12
        };
        return monthMap[s] || null;
    }

    function normalizeTime(value) {
        if (
            value === null ||
            value === undefined ||
            value === ""
        ) {
            return "";
        }
        const s = String(value).trim();
        if (!s) {
            return "";
        }
        /* 08:30 */
        const timeMatch = s.match(/^(\d{1,2}):(\d{2})$/);
        if (timeMatch) {
            return `${String(timeMatch[1]).padStart(2, "0")}:${timeMatch[2]}`;
        }
        /* 08.30 */
        const dotMatch = s.match(/^(\d{1,2})\.(\d{2})$/);
        if (dotMatch) {
            return `${String(dotMatch[1]).padStart(2, "0")}:${dotMatch[2]}`;
        }
        /* Excel / Date */
        const date = new Date(`1970-01-01 ${s}`);
        if (!Number.isNaN(date.getTime())) {
            return `${String(date.getHours()).padStart(2, "0")}:${String(date.getMinutes()).padStart(2, "0")}`;
        }
        return s;
    }

    function parseDateValue(value, year, month) {
        if (
            value === null ||
            value === undefined ||
            value === ""
        ) {
            return "";
        }
        const s = String(value).trim();
        /* YYYY-MM-DD */
        const isoMatch = s.match(
            /(\d{4})[-/](\d{1,2})[-/](\d{1,2})/
        );
        if (isoMatch) {
            return `${isoMatch[1]}-${String(isoMatch[2]).padStart(2, "0")}-${String(isoMatch[3]).padStart(2, "0")}`;
        }
        /* วันที่เป็นเลขอย่างเดียว เช่น 3 */
        if (
            /^\d{1,2}$/.test(s) &&
            year &&
            month
        ) {
            return `${year}-${String(month).padStart(2, "0")}-${String(s).padStart(2, "0")}`;
        }
        const d = new Date(s);
        if (!Number.isNaN(d.getTime())) {
            return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
        }
        return s;
    }

    function normalizeFlight(value) {
        const s = safeText(value).toLowerCase();
        if (!s) {
            return "International";
        }
        if (s.includes("dom")) {
            return "Domestic";
        }
        if (s.includes("int")) {
            return "International";
        }
        return s.charAt(0).toUpperCase() + s.slice(1);
    }

    function normalizeDirection(value) {
        const s = safeText(value).toLowerCase();
        if (!s) {
            return "Arrival";
        }
        if (s.includes("dep")) {
            return "Departure";
        }
        if (s.includes("arr")) {
            return "Arrival";
        }
        if (s.includes("out")) {
            return "Departure";
        }
        if (s.includes("in")) {
            return "Arrival";
        }
        return s.charAt(0).toUpperCase() + s.slice(1);
    }

    function parseMatrixSheet(sheet) {
        const matrix = XLSX.utils.sheet_to_json(sheet, {
            header: 1,
            defval: "",
            raw: true
        });
        if (!matrix || matrix.length < 3) return [];
        // Sheet names are the source of month/year for this workbook.
        const sheetName = safeText(sheet.__sheetName || "");
        const m = sheetName.match(/^(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{2})$/i);
        const month = m ? parseMonthValue(m[1]) : null;
        const year = m ? 2000 + Number(m[2]) : new Date().getFullYear();
        // Row 3 (index 2) contains the day numbers.
        const dayRow = matrix[2] || [];
        const rows = [];
        let flight = safeText(matrix[0]?.[0]) === "International"
            ? "International"
            : safeText(matrix[0]?.[0]) === "Domestic"
                ? "Domestic"
                : "";
        let direction = "";
        const serviceNames = {
            "Check-In": "Check-In",
            "In-Line\nScreening": "In-Line Screening",
            "Security\nScreening": "Security Screening",
            "Immigration": "Immigration",
            "Baggage\nClaim": "Baggage Claim",
            "Taxi": "Taxi"
        };
        const normalizeCellTime = (value) => {
            if (value instanceof Date) {
                return `${String(value.getHours()).padStart(2, "0")}:${String(value.getMinutes()).padStart(2, "0")}`;
            }
            if (typeof value === "number") {
                // Excel time stored as fraction of a day.
                if (value >= 0 && value < 1) {
                    const totalMinutes = Math.round(value * 24 * 60);
                    return `${String(Math.floor(totalMinutes / 60) % 24).padStart(2, "0")}:${String(totalMinutes % 60).padStart(2, "0")}`;
                }
            }
            return normalizeTime(value);
        };
        for (let i = 3; i < matrix.length; i++) {
            const row = matrix[i] || [];
            const first = safeText(row[0]).replace(/\r?\n/g, "\n");
            if (first === "International" || first === "Domestic") {
                flight = first;
                direction = "";
                continue;
            }
            if (first === "Departure" || first === "Arrival") {
                direction = first;
                continue;
            }
            const service = serviceNames[first];
            if (!service || !direction || !flight) continue;
            // Each service occupies four rows: Start, Finish, Waiting, Passenger.
            const startRow = row;
            const finishRow = matrix[i + 1] || [];
            const waitingRow = matrix[i + 2] || [];
            const passengerRow = matrix[i + 3] || [];
            for (let col = 2; col < dayRow.length; col++) {
                const day = Number(dayRow[col]);
                if (!Number.isInteger(day) || day < 1 || day > 31) continue;
                const startRaw = startRow[col];
                const finishRaw = finishRow[col];
                const waitingRaw = waitingRow[col];
                // Baggage Claim has no separate Passenger row in the matrix.
                // For International Arrival, the Passenger count is recorded on the
                // preceding Immigration block for the same date/column.
                const hasPassengerRow = safeText(passengerRow[1]) === "Passenger";
                let passengerRaw = hasPassengerRow ? passengerRow[col] : "";
                if (!hasPassengerRow && service === "Baggage Claim" && direction === "Arrival" && flight === "International") {
                    // Immigration Passenger row is exactly 1 row before Baggage Claim Start.
                    const linkedPassengerRow = matrix[i - 1] || [];
                    if (safeText(linkedPassengerRow[1]) === "Passenger") passengerRaw = linkedPassengerRow[col];
                }
                const hasAny =
                    (startRaw !== "" && startRaw !== null && startRaw !== undefined) ||
                    (finishRaw !== "" && finishRaw !== null && finishRaw !== undefined) ||
                    (passengerRaw !== "" && passengerRaw !== null && passengerRaw !== undefined);
                if (!hasAny) continue;
                const start = normalizeCellTime(startRaw);
                const finish = normalizeCellTime(finishRaw);
                const waitingValue = excelDurationToMinutes(waitingRaw);
                const waiting = waitingValue !== null ? waitingValue : calculateWaitingMinutes(start, finish, 0);
                const passenger = toNumeric(passengerRaw);

                rows.push({
                    year,
                    month: month || 1,
                    monthLabel: monthNames[month] || "เดือน",
                    date: `${year}-${String(month || 1).padStart(2, "0")}-${String(day).padStart(2, "0")}`,
                    flight,
                    direction,
                    service,
                    start,
                    finish,
                    waiting,
                    passenger
                });
            }
        }
        return rows;
    }

    function parseGenericSheet(sheet) {
        const json =
            XLSX.utils.sheet_to_json(
                sheet,
                {
                    defval: "",
                    raw: false
                }
            );

        if (
            !json ||
            !json.length
        ) {
            return [];
        }

        const rows = [];

        for (const item of json) {
            const normalized =
                normalizeRow(item);
            if (normalized) {
                rows.push(normalized);
            }
        }

        return rows;
    }

    function extractRowsFromWorkbook(workbook) {
        const rows = [];

        for (
            const sheetName of workbook.SheetNames
        ) {
            const sheet =
                workbook.Sheets[sheetName];
            sheet.__sheetName = sheetName;

            const matrixRows = parseMatrixSheet(sheet);
            if (matrixRows.length) {
                rows.push(...matrixRows);
                continue;
            }
            const byJson = parseGenericSheet(sheet);
            if (byJson.length) rows.push(...byJson);
        }

        return rows;
    }

    // Convert Excel time fraction (e.g. 35 sec = 35/86400 ≈ 0.000405) to decimal minutes.
    function excelDurationToMinutes(value) {
        if (value === "" || value === null || value === undefined) return null;
        if (typeof value === "number") {
            if (value > 0 && value < 1) return value * 24 * 60; // Excel fraction of day
            if (value >= 1) return value; // already in minutes
            return null;
        }
        const s = String(value).trim();
        const mmss = s.match(/^(\d{1,3}):(\d{2})$/);
        if (mmss) return Number(mmss[1]) + Number(mmss[2]) / 60; // "M:SS" string
        const n = toNumeric(s);
        return Number.isFinite(n) && n >= 0 ? n : null;
    }
    function getServiceList(direction, flight) {
        if (direction && direction !== "all" && flight && flight !== "all") {
            const key = `${direction}-${flight}`;
            if (SERVICES[key]) return SERVICES[key];
        }
        return SERVICES[direction] || ALL_SERVICES;
    }

    function readFilters() {
        return Object.fromEntries(FILTER_IDS.map(id => [id, elements[id].value]));
    }

    function populateSelect(select, options) {
        const current = select.value;
        select.innerHTML = '<option value="all">ทั้งหมด</option>' + options.map(value =>
            `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`
        ).join("");
        const hasCurrent = Array.from(select.options).some(option => option.value === current);
        select.value = hasCurrent ? current : "all";
    }

    function populateFilters() {
        populateSelect(elements.year, CONFIG.dashboardYears);
        updateServiceOptions();
    }

    function updateServiceOptions() {
        populateSelect(elements.service, getServiceList(elements.direction.value, elements.flight.value));
    }

    function updateDateOptions() {
        const { year, month } = readFilters();
        const maxDay = month === "all"
            ? 31
            : new Date(year === "all" ? 2024 : Number(year), Number(month), 0).getDate();
        populateSelect(elements.date, Array.from({ length: maxDay }, (_, index) => String(index + 1)));
        renderPicker("date");
    }

    function renderPicker(id) {
        const select = elements[id];
        const menu = elements[`${id}PickerMenu`];
        menu.innerHTML = Array.from(select.options).map(option => `
            <div class="${id}-picker-option ${option.value === select.value ? "selected" : ""}"
                 data-value="${escapeHtml(option.value)}">${escapeHtml(option.textContent)}</div>
        `).join("");
        elements[`${id}PickerButton`].textContent = select.options[select.selectedIndex]?.textContent || "ทั้งหมด";
    }

    function closePickerMenus(exceptId) {
        for (const id of ["month", "date"]) {
            if (id !== exceptId) elements[`${id}PickerMenu`].classList.remove("open");
        }
    }

    function setupPicker(id) {
        const picker = elements[`${id}Picker`];
        const menu = elements[`${id}PickerMenu`];
        elements[`${id}PickerButton`].addEventListener("click", event => {
            event.stopPropagation();
            closePickerMenus(id);
            menu.classList.toggle("open");
        });
        // One delegated listener survives each menu rebuild.
        menu.addEventListener("click", event => {
            const option = event.target.closest(`.${id}-picker-option`);
            if (!option || !menu.contains(option)) return;
            elements[id].value = option.dataset.value;
            closePickerMenus();
            elements[id].dispatchEvent(new Event("change"));
        });
        document.addEventListener("click", event => {
            if (!picker.contains(event.target)) menu.classList.remove("open");
        });
        renderPicker(id);
    }

    function getFilteredRows() {
        const filters = readFilters();
        return state.rows.filter(row => {
            if (filters.year !== "all" && String(row.year) !== filters.year) return false;
            if (filters.month !== "all" && String(row.month) !== filters.month) return false;
            if (filters.date !== "all" && String(row.date).slice(-2) !== filters.date.padStart(2, "0")) return false;
            return ["flight", "direction", "service"].every(id =>
                filters[id] === "all" || row[id] === filters[id]
            );
        });
    }

    function resetFilters() {
        FILTER_IDS.forEach(id => { elements[id].value = "all"; });
        updateServiceOptions();
        updateDateOptions();
        renderPicker("month");
        closePickerMenus();
        updateDashboard();
    }

    // Shared summaries and dashboard rendering.
    function summarizeRows(rows) {
        let passengers = 0;
        let waitingTotal = 0;
        let waitingCount = 0;
        for (const row of rows) {
            passengers += toNumeric(row.passenger);
            const waiting = toNumeric(row.waiting);
            if (Number.isFinite(waiting) && waiting > 0) {
                waitingTotal += waiting;
                waitingCount++;
            }
        }
        return {
            records: rows.length,
            passengers,
            averageWaiting: waitingCount ? waitingTotal / waitingCount : 0
        };
    }

    function updateKpi(rows) {
        const summary = summarizeRows(rows);
        elements.passenger.textContent = formatNumber(summary.passengers);
        elements.waiting.textContent = formatWaitingDuration(summary.averageWaiting);
        elements.records.textContent = formatNumber(summary.records);
    }

    function updateCards(rows) {
        const { direction, service: selectedService, flight } = readFilters();
        const services = getServiceList(direction, flight).filter(service =>
            selectedService === "all" || service === selectedService
        );
        const grouped = new Map(services.map(service => [service, []]));
        rows.forEach(row => grouped.get(row.service)?.push(row));
        elements.cards.innerHTML = services.map(service => {
            const summary = summarizeRows(grouped.get(service));
            const passengers = summary.passengers > 0 ? formatNumber(summary.passengers) : "-";
            return `
                <div class="card">
                    <h3>${escapeHtml(service)}</h3>
                    <div class="muted">${summary.records} รายการ</div>
                    <div class="grid2">
                        <div class="metric"><small>Passenger</small><b>${passengers}</b></div>
                        <div class="metric"><small>Avg. Waiting</small><b>${formatWaitingDuration(summary.averageWaiting)}</b></div>
                    </div>
                </div>
            `;
        }).join("");
    }

    function updateTable(rows) {
        elements.rowCount.textContent = `${rows.length} รายการ`;
        if (!rows.length) {
            elements.tableBody.innerHTML = '<tr><td colspan="8">ไม่มีข้อมูล</td></tr>';
            return;
        }
        elements.tableBody.innerHTML = rows.map(row => {
            const cells = [
                row.year, monthNames[row.month] || row.month, formatDateThai(row.date),
                row.flight, row.direction, row.service,
                formatWaitingDuration(row.waiting), formatNumber(row.passenger)
            ];
            return `<tr>${cells.map(value => `<td>${escapeHtml(value)}</td>`).join("")}</tr>`;
        }).join("");
    }

    function updateExcelButton() {
        const { year, month } = readFilters();
        let url = CONFIG.excelUrl;
        let title = "เปิดไฟล์ Excel";
        let subtitle = "ดูข้อมูลต้นฉบับ";
        if (year !== "all" && month !== "all") {
            const monthNumber = Number(month);
            const sheetName = `${CONFIG.monthNamesEn[monthNumber]} ${year.slice(-2)}`;
            url += `&activeCell=${encodeURIComponent(`'${sheetName}'!A1`)}`;
            title = `เปิด Excel เดือน ${monthNames[monthNumber]} ${year}`;
            subtitle = `เปิดข้อมูลชีต ${sheetName}`;
        } else if (year !== "all") {
            title = `เปิดไฟล์ Excel ปี ${year}`;
        }
        elements.excelButton.href = url;
        elements.excelButtonTitle.textContent = title;
        elements.excelButtonSub.textContent = subtitle;
    }

    function updateDashboard() {
        const rows = getFilteredRows();
        updateExcelButton();
        updateKpi(rows);
        updateCards(rows);
        updateTable(rows);
    }

    // Data loading and startup.
    async function loadExcel() {
        try {
            const response = await fetch(CONFIG.excelFile, { cache: "no-store" });
            if (!response.ok) throw new Error(`HTTP ${response.status}`);
            const buffer = await response.arrayBuffer();
            if (!buffer.byteLength) throw new Error("ไฟล์ Excel ว่าง");
            const rows = extractRowsFromWorkbook(XLSX.read(buffer, { type: "array" }));
            if (!rows.length) throw new Error("ไม่พบข้อมูลที่ใช้งานได้");
            state.rows = rows;
            elements.status.textContent = `โหลดข้อมูลจาก ${CONFIG.excelFile} สำเร็จ: ${rows.length} รายการ`;
        } catch (error) {
            console.error("Excel loading error:", error);
            elements.status.textContent = `ไม่สามารถโหลดข้อมูลได้: ${error.message}`;
            elements.status.className = "status error";
            return;
        }
        elements.mainDashboard.style.display = "block";
        populateFilters();
        updateDashboard();
    }

    function initializeDashboard() {
        document.querySelectorAll("[id]").forEach(element => { elements[element.id] = element; });
        FILTER_IDS.forEach(id => {
            elements[id].addEventListener("change", () => {
                if (id === "year" || id === "month") updateDateOptions();
                if (id === "month" || id === "date") renderPicker(id);
                if (id === "direction" || id === "flight") updateServiceOptions();
                updateDashboard();
            });
        });
        elements.resetButton.addEventListener("click", resetFilters);
        setupPicker("month");
        setupPicker("date");
        updateDateOptions();
        loadExcel();
    }

    document.addEventListener("DOMContentLoaded", initializeDashboard, { once: true });
})();
