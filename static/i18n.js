(function (root, factory) {
    const api = factory();
    if (typeof module === 'object' && module.exports) {
        module.exports = api;
    }
    root.I18N = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
    const UI = {
        en: {
            pageTitle: 'Metropolis Cost Allocation Dashboard',
            billingMonth: 'Billing Month',
            monthFormat: 'YYYY-MM',
            setMonth: 'Set Month',
            navigation: 'Navigation',
            intake: 'Intake',
            consumption: 'Consumption',
            toggleTheme: 'Toggle Theme',
            ariaLanguage: 'Language',
            matched: 'Matched',
            missing: 'Missing',
            wrongMonth: 'Wrong Month',
            duplicates: 'Duplicates',
            gateStatus: 'Gate Status',
            blocked: 'Blocked',
            passed: 'Passed',
            uploadBills: 'Upload Bills',
            selectUpload: 'Select and upload PDF files',
            selectPdf: 'Select PDF files (up to 15)',
            uploadParse: 'Upload & Parse Bills',
            pleaseSelectPdf: 'Please select PDF files',
            uploading: 'Uploading and parsing...',
            lastUpload: 'Last upload: {time}',
            parsedBills: 'Parsed {n} bills',
            unknownError: 'Unknown error',
            expectedAccounts: '15 Expected Accounts',
            accountsSubtitle: 'Metropolis CLP accounts with live status',
            roleRetailChillers: 'Retail chillers → AC/SW',
            roleOfficeMulti: 'Multi-centre (office)',
            roleOfficeChiller: 'Office chiller → AO/OC',
            roleCarpark: 'Multi-centre (carpark)',
            roleDc: 'DC 100%',
            roleSw: 'SW 100%',
            roleFitRetail: 'FiT (retail)',
            roleFit: 'FiT',
            roleC: 'C 100%',
            roleSa: 'SA 100%',
            roleFoodCourt: 'Food Court → FC 100%',
            meterLog: 'Meter Log',
            meterSubtitle: 'Check-meter 6681757 (散熱水泵電 / Sea Water Pump)',
            save: 'Save',
            downloadExcel: 'Download Excel',
            previousReading: 'Previous Reading',
            presentReading: 'Present Reading',
            readDate: 'Read Date (optional)',
            examplePrevious: 'e.g., 95966.6',
            examplePresent: 'e.g., 96323.1',
            pleaseEnterReadings: 'Please enter both readings',
            saving: 'Saving...',
            savedDelta: 'Saved: Δ = {n} kWh',
            gateValidation: 'Gate Validation',
            gatePrompt: 'Set month and parse bills to validate gate conditions',
            continuePart2: 'Continue to Part 2',
            gatePassedTitle: 'Gate PASSED',
            gateReady: 'All conditions met. Ready to continue to Part 2.',
            gateSummary: '{matched}/15 matched, {missing} missing',
            statusMatched: 'Matched',
            statusMissing: 'Missing',
            statusDuplicate: 'Duplicate',
            statusWrongMonth: 'Wrong Month',
            statusUnrecognised: 'Unrecognised',
            processingMonthEnd: 'Processing month-end...',
            copyingWorkbooks: "Copying previous month's workbooks and updating with current data...",
            monthEndComplete: 'Month-End Complete',
            updatedWorkbooks: 'Updated workbooks:',
            monthEndFailed: 'Month-End Failed',
            monthEndFailedSummary: 'Month-end processing failed.',
            monthEndError: 'Error processing month-end.',
            downloads: 'Downloads',
            downloadsSubtitle: 'Workbooks written for the selected month, and the history file',
            month: 'Month',
            downloadAllocation: 'Download Cost Allocation',
            downloadSheet: 'Download Cost Sheet',
            downloadHistory: 'Download history',
            history: 'History',
            historySubtitle: 'One button opens that item across the saved months',
            historyNote: "Bill kWh and bill charges stay in the history workbook. Choose electricity bills, meters, or cost centres. Only that set of buttons is shown. Nothing is charted until a button is clicked. A cost centre's kWh is each bill's saved kWh times that month's live percent. Its cost is the allocation amount. A meter's second series is a calculated share of the account charge, not an amount the bill printed.",
            electricityBills: 'Electricity bills',
            meters: 'Meters',
            costCentres: 'Cost centres',
            ariaWhichButtons: 'Which buttons to show',
            chartPlaceholder: "Choose a button to open that item's history.",
            kwh: 'kWh',
            cost: 'Cost',
            proportionalCharge: 'Proportional charge',
            nothingSaved: 'Nothing saved for this choice yet.',
            noProcessedMonth: 'No processed month',
            couldNotLoadButtons: 'Could not load buttons',
            noHistoryItem: 'No history for that item',
            noteZeroKwh: 'Account kWh is zero, so the proportional charge is blank.',
            noCostStored: 'No cost is stored for this item.',
            nothingToChart: 'Nothing to chart for this item.',
            notPrinted: 'not printed',
            kwhChartAria: 'kWh chart',
            costChartAria: 'Cost chart',
            networkError: 'Error: {message}',
        },
        zh: {
            pageTitle: 'Metropolis 電費成本分攤',
            billingMonth: '帳單月份',
            monthFormat: 'YYYY-MM',
            setMonth: '設定月份',
            navigation: '導覽',
            intake: '匯入',
            consumption: '用電量',
            toggleTheme: '切換主題',
            ariaLanguage: '語言',
            matched: '已相符',
            missing: '缺漏',
            wrongMonth: '月份不符',
            duplicates: '重複',
            gateStatus: '閘口狀態',
            blocked: '未通過',
            passed: '已通過',
            uploadBills: '上載電費單',
            selectUpload: '選擇並上載 PDF 檔案',
            selectPdf: '選擇 PDF 檔案（最多 15 份）',
            uploadParse: '上載並讀取電費單',
            pleaseSelectPdf: '請選擇 PDF 檔案',
            uploading: '正在上載並讀取…',
            lastUpload: '上次上載：{time}',
            parsedBills: '已讀取 {n} 份電費單',
            unknownError: '未知錯誤',
            expectedAccounts: '15 個預期帳戶',
            accountsSubtitle: 'Metropolis 中電帳戶及即時狀態',
            roleRetailChillers: '商場冷氣機 → AC/SW',
            roleOfficeMulti: '多個成本中心（辦公室）',
            roleOfficeChiller: '辦公室冷氣機 → AO/OC',
            roleCarpark: '多個成本中心（停車場）',
            roleDc: 'DC 100%',
            roleSw: 'SW 100%',
            roleFitRetail: '上網電價（商場）',
            roleFit: '上網電價',
            roleC: 'C 100%',
            roleSa: 'SA 100%',
            roleFoodCourt: '美食廣場 → FC 100%',
            meterLog: '電錶紀錄',
            meterSubtitle: '核對電錶 6681757（散熱水泵電）',
            save: '儲存',
            downloadExcel: '下載 Excel',
            previousReading: '上期讀數',
            presentReading: '今期讀數',
            readDate: '抄錶日期（可選）',
            examplePrevious: '例如 95966.6',
            examplePresent: '例如 96323.1',
            pleaseEnterReadings: '請輸入上期及今期讀數',
            saving: '正在儲存…',
            savedDelta: '已儲存：差額 = {n} kWh',
            gateValidation: '閘口檢查',
            gatePrompt: '請先設定月份並讀取電費單，以檢查閘口條件',
            continuePart2: '繼續月結',
            gatePassedTitle: '閘口已通過',
            gateReady: '所有條件已符合，可繼續月結。',
            gateSummary: '{matched}/15 已相符，缺漏 {missing}',
            statusMatched: '已相符',
            statusMissing: '缺漏',
            statusDuplicate: '重複',
            statusWrongMonth: '月份不符',
            statusUnrecognised: '未能識別',
            processingMonthEnd: '正在處理月結…',
            copyingWorkbooks: '正在複製上月工作簿，並以本月資料更新…',
            monthEndComplete: '月結完成',
            updatedWorkbooks: '已更新的工作簿：',
            monthEndFailed: '月結失敗',
            monthEndFailedSummary: '月結處理失敗。',
            monthEndError: '處理月結時發生錯誤。',
            downloads: '下載',
            downloadsSubtitle: '所選月份的工作簿，以及用電紀錄檔',
            month: '月份',
            downloadAllocation: '下載成本分攤',
            downloadSheet: '下載成本表',
            downloadHistory: '下載用電紀錄',
            history: '紀錄',
            historySubtitle: '按一下按鈕，查看該項目在已儲存月份的走勢',
            historyNote: '電費單的用電量及費用保存在用電紀錄工作簿。可選擇電費單、電錶或成本中心，頁面每次只顯示其中一組按鈕。按下按鈕後才會繪圖。成本中心的用電量是各電費單已儲存用電量乘以該月的實際百分比，費用是分攤金額。電錶的第二個系列是按帳戶費用及用電量比例計算的分攤，並非電費單上列印的金額。',
            electricityBills: '電費單',
            meters: '電錶',
            costCentres: '成本中心',
            ariaWhichButtons: '顯示哪一組按鈕',
            chartPlaceholder: '請選擇按鈕以開啟該項目的紀錄。',
            kwh: '用電量',
            cost: '費用',
            proportionalCharge: '按比例分攤費用',
            nothingSaved: '此選項尚未有儲存紀錄。',
            noProcessedMonth: '未有已處理月份',
            couldNotLoadButtons: '無法載入按鈕',
            noHistoryItem: '此項目沒有已儲存紀錄',
            noteZeroKwh: '帳戶用電量為零，因此按比例分攤費用留空。',
            noCostStored: '此項目沒有儲存費用。',
            nothingToChart: '此項目沒有可繪製的資料。',
            notPrinted: '未有列印金額',
            kwhChartAria: '用電量圖',
            costChartAria: '費用圖',
            networkError: '錯誤：{message}',
        },
    };

    const SERVER = [
        { en: 'Please set billing month first', zh: '請先設定帳單月份' },
        { en: 'No files uploaded', zh: '沒有上載檔案' },
        { en: 'Present reading must be >= previous reading', zh: '今期讀數必須大於或等於上期讀數' },
        { en: 'No billing month set', zh: '尚未設定帳單月份' },
        { en: 'Billing month not set', zh: '尚未設定帳單月份' },
        { en: 'Meter log incomplete or invalid (present must be ≥ previous)', zh: '電錶紀錄未完成或無效（今期讀數必須大於或等於上期）' },
        { en: 'No bills parsed', zh: '尚未讀取電費單' },
        { en: 'Cost Allocation master not found. Checked previous/current month folders.', zh: '找不到成本分攤主檔。已檢查上月及本月資料夾。' },
        { en: 'No allocation rules found in master workbook AC DEPT sheet', zh: '主檔工作簿的 AC DEPT 工作表沒有分攤規則' },
        { en: 'Month must be YYYY-MM', zh: '月份必須為 YYYY-MM' },
        { en: 'kind must be bill, meter, or centre', zh: '種類必須為電費單、電錶或成本中心' },
        { en: 'item is required', zh: '必須指定項目' },
        { en: 'No saved history for that item', zh: '此項目沒有已儲存紀錄' },
        { en: 'No electricity history yet. Process a month first.', zh: '尚未有用電紀錄。請先處理一個月份。' },
        { en: 'kind must be cost-allocation or cost-sheet', zh: '種類必須為成本分攤或成本表' },
        { en: 'Invalid month format. Use YYYY-MM (e.g., 2025-04)', zh: '月份格式不正確。請使用 YYYY-MM（例如 2025-04）' },
        { en: 'Account kWh is zero, so the proportional charge is blank.', zh: '帳戶用電量為零，因此按比例分攤費用留空。' },
        { en: 'Could not load buttons', zh: '無法載入按鈕' },
        { en: 'No history for that item', zh: '此項目沒有已儲存紀錄' },
        { pattern: '^Folder not found: ([\\s\\S]+)\\nCreate the folder and place CLP bill PDFs inside\\.$', zh: '找不到資料夾：{1}\n請建立資料夾，並放入中電電費單 PDF。' },
        { pattern: '^No PDF files found in (.+)$', zh: '在 {1} 找不到 PDF 檔案' },
        { prefix: 'Error scanning folder: ', zh: '掃描資料夾時發生錯誤：{1}' },
        { prefix: 'Error uploading bills: ', zh: '上載電費單時發生錯誤：{1}' },
        { prefix: 'Error setting meter log: ', zh: '儲存電錶紀錄時發生錯誤：{1}' },
        { prefix: 'Error generating meter log Excel: ', zh: '產生電錶紀錄 Excel 時發生錯誤：{1}' },
        { prefix: 'Error processing month-end: ', zh: '處理月結時發生錯誤：{1}' },
        { prefix: 'Gate must pass before processing month-end. Errors: ', zh: '必須先通過閘口才可處理月結。錯誤：{1}' },
        { pattern: '^Could not compute previous month from (\\d{4}-\\d{2})$', zh: '無法從 {1} 計算上一個月' },
        { pattern: '^No dashboard report for (\\d{4}-\\d{2})\\. Process that month first\\.$', zh: '尚未有 {1} 的儀表板紀錄。請先處理該月。' },
        { pattern: '^Month must be YYYY-MM: (.+)$', zh: '月份必須為 YYYY-MM：{1}' },
        { pattern: '^No cost-allocation workbook for (\\d{4}-\\d{2})$', zh: '找不到 {1} 的成本分攤工作簿' },
        { pattern: '^No cost-sheet workbook for (\\d{4}-\\d{2})$', zh: '找不到 {1} 的成本表工作簿' },
        { pattern: '^Month-end processed: copied from (\\d{4}-\\d{2}), updated for (\\d{4}-\\d{2})$', zh: '月結已處理：已由 {1} 複製，並更新為 {2}' },
        { pattern: '^Missing (\\d+) accounts: (.+)$', zh: '缺漏 {1} 個帳戶：{2}' },
        { pattern: '^Duplicate accounts: (.+)$', zh: '重複帳戶：{1}' },
        { pattern: '^Wrong month: (.+)$', zh: '月份不符：{1}' },
        { pattern: '^Unrecognised accounts \\(not in Metropolis 15\\): (.+)$', zh: '未能識別的帳戶（不在 Metropolis 15 個帳戶內）：{1}' },
        { suffix: ': Not a PDF', zh: '{1}：不是 PDF' },
        { suffix: ': Failed to parse (no KWH)', zh: '{1}：無法讀取（沒有用電量）' },
    ];

    let active = 'en';

    function savedLanguage(storage) {
        return storage && storage.getItem('dashboard-lang') === 'zh' ? 'zh' : 'en';
    }

    function rememberLanguage(storage, lang) {
        const chosen = lang === 'zh' ? 'zh' : 'en';
        storage.setItem('dashboard-lang', chosen);
        return chosen;
    }

    function setActiveLanguage(lang) {
        active = lang === 'zh' ? 'zh' : 'en';
        return active;
    }

    function activeLanguage() {
        return active;
    }

    function fill(text, vars) {
        let out = text;
        Object.keys(vars || {}).forEach((name) => {
            out = out.split('{' + name + '}').join(String(vars[name]));
        });
        return out;
    }

    function t(key, vars) {
        const table = UI[active] || UI.en;
        const text = table[key] != null ? table[key] : (UI.en[key] != null ? UI.en[key] : key);
        return fill(text, vars);
    }

    function applyRule(text, rule) {
        if (rule.en && text === rule.en) return rule.zh;
        if (rule.prefix && text.startsWith(rule.prefix)) {
            return fill(rule.zh, { 1: text.slice(rule.prefix.length) });
        }
        if (rule.suffix && text.endsWith(rule.suffix)) {
            return fill(rule.zh, { 1: text.slice(0, text.length - rule.suffix.length) });
        }
        if (rule.pattern) {
            const match = text.match(new RegExp(rule.pattern));
            if (match) {
                return rule.zh.replace(/\{(\d+)\}/g, (_, index) => match[Number(index)] ?? '');
            }
        }
        return null;
    }

    function translateFragments(text) {
        const exact = SERVER.filter((rule) => rule.en).slice().sort((a, b) => b.en.length - a.en.length);
        let out = text;
        exact.forEach((rule) => {
            out = out.split(rule.en).join(rule.zh);
        });
        return out;
    }

    function translateServer(message) {
        if (message == null) return '';
        const text = String(message);
        if (active !== 'zh') return text;
        if (text.includes(' • ')) {
            return text.split(' • ').map((part) => translateServer(part)).join(' • ');
        }
        for (let index = 0; index < SERVER.length; index += 1) {
            const rendered = applyRule(text, SERVER[index]);
            if (rendered != null) return translateFragments(rendered);
        }
        return translateFragments(text);
    }

    return {
        UI: UI,
        SERVER: SERVER,
        savedLanguage: savedLanguage,
        rememberLanguage: rememberLanguage,
        setActiveLanguage: setActiveLanguage,
        activeLanguage: activeLanguage,
        t: t,
        translateServer: translateServer,
    };
});
