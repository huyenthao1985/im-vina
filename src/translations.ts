export interface TranslationDict {
  marketingInsights: string;
  mainTitleUpload: string;
  mainTitleDash: string;
  subtitleUpload: string;
  subtitleDash: string;
  
  // Upload Zone
  dragDrop: string;
  orClickBrowse: string;
  selectExcelBtn: string;
  sampleDataBtn: string;
  salesDataBtn: string;
  formatsHint: string;
  browserProcessOnly: string;
  invalidFileError: string;
  loadingError: string;
  
  // Filters
  fromYear: string;
  toYear: string;
  partner: string;
  customer: string;
  resetBtn: string;
  loadExcelBtn: string;
  filtersApplied: string;
  rowsPill: string;
  allOption: string;
  
  // KPIs
  totalSales: string;
  totalShipment: string;
  totalProduction: string;
  conversionRatio: string;
  salesGrowthYoY: string;
  insufficientYears: string;
  growthComparison: string;
  
  // Charts
  chart1Title: string;
  chart1Sub: string;
  chartMonthTitle: string;
  chartMonthSub: string;
  chart1BarName: string;
  chart1LineName: string;
  
  chart2Title: string;
  chart2Sub: string;
  
  chart3Title: string;
  chart3Sub: string;
  
  chart4Title: string;
  chart4Sub: string;
  
  chart5Title: string;
  chart5Sub: string;
  
  chart6Title: string;
  chart6Sub: string;
  
  chart7Title: string;
  chart7Sub: string;

  chartModelAllYearTitle: string;
  chartModelAllYearSub: string;
  chartModelFilteredTitle: string;
  chartModelFilteredSub: string;
  chartCustomerStackedTitle: string;
  chartCustomerStackedSub: string;
  chartCustomerDonutTitle: string;
  chartCustomerDonutSub: string;
  
  // Funnel
  prodLabel: string;
  shipLabel: string;
  salesLabel: string;
  
  // Table
  tableTitle: string;
  colModel: string;
  colPartner: string;
  colCustomer: string;
  colType: string;
  colProd: string;
  colShip: string;
  colSales: string;
  colRatio: string;
  tableEmpty: string;
  dbTitle: string;
  dbSearchPlace: string;
  colDivision: string;
  colYear: string;
  colMonth: string;
  colValue: string;
  showingRows: string;
  rowsPerPage: string;
}

export const translations: Record<'vi' | 'en' | 'ko', TranslationDict> = {
  vi: {
    marketingInsights: 'Marketing Insights',
    mainTitleUpload: 'Tổng quan Sản xuất – Xuất hàng – Doanh số',
    mainTitleDash: 'Tổng quan Sản xuất – Xuất hàng – Doanh số',
    subtitleUpload: 'Tải lên tệp Excel để phân tích theo Model, Đối tác, Khách hàng và thời gian',
    subtitleDash: 'Phân tích theo Model, Đối tác, Khách hàng và thời gian',
    dragDrop: 'Kéo thả file Excel vào đây',
    orClickBrowse: 'hoặc click để duyệt từ máy tính của bạn',
    selectExcelBtn: 'Chọn File Excel (.xlsx)',
    sampleDataBtn: 'Dữ liệu mẫu (Marketing)',
    salesDataBtn: 'Dữ liệu Sales (Test 2.xlsx)',
    formatsHint: 'Hỗ trợ định dạng .xlsx và .xls',
    browserProcessOnly: 'Dữ liệu được xử lý trực tiếp trên trình duyệt của bạn, không tải lên máy chủ nào.',
    invalidFileError: 'Chỉ chấp nhận tệp .xlsx hoặc .xls',
    loadingError: 'Không thể đọc tệp. Tệp có thể bị hỏng hoặc định dạng không hỗ trợ.',
    fromYear: 'Từ năm',
    toYear: 'Đến năm',
    partner: 'Custom',
    customer: 'Khách hàng',
    resetBtn: 'Đặt lại',
    loadExcelBtn: 'Tải Excel khác',
    filtersApplied: 'bộ lọc đang áp dụng',
    rowsPill: 'dòng dữ liệu',
    allOption: 'Tất cả',
    totalSales: 'Tổng Doanh số',
    totalShipment: 'Tổng Xuất hàng',
    totalProduction: 'Tổng Sản xuất',
    conversionRatio: 'Tỷ lệ Xuất hàng/Sản xuất',
    salesGrowthYoY: 'Tăng trưởng Doanh số (YoY)',
    insufficientYears: 'Chưa đủ 2 năm trọn vẹn',
    growthComparison: 'so với',
    chart1Title: 'Xuất hàng & Doanh số theo Năm',
    chart1Sub: 'Cột: Xuất hàng (K) · Đường: Doanh số (K$) · Nét đứt: Tăng trưởng (YoY)',
    chartMonthTitle: 'Xuất hàng & Doanh số theo Tháng',
    chartMonthSub: 'Cột: Xuất hàng (K) · Đường: Doanh số (K$) · Nét đứt: Tăng trưởng (MoM)',
    chart1BarName: 'Xuất hàng (K)',
    chart1LineName: 'Doanh số (K$)',
    chart2Title: 'Phễu Sản xuất → Xuất hàng → Doanh số',
    chart2Sub: 'Tổng theo phạm vi đang lọc',
    chart3Title: 'Tỷ trọng theo Đối tác',
    chart3Sub: 'Theo Doanh số',
    chart4Title: 'Tỷ lệ Xuất hàng / Sản xuất',
    chart4Sub: 'Hiệu suất chuyển đổi',
    chart5Title: 'Top 10 Model theo Doanh số',
    chart5Sub: 'K$ · sắp xếp giảm dần',
    chart6Title: 'Sản xuất theo Loại (TYPE)',
    chart6Sub: 'Tổng sản lượng (K)',
    chart7Title: 'Tỷ trọng theo Khách hàng',
    chart7Sub: 'Theo Doanh số',
    chartModelAllYearTitle: 'Top Model - ALL YEAR',
    chartModelAllYearSub: 'Cột: Top Model doanh số (K$) · Xanh lam: Bán chạy nhất · Vàng: Khác',
    chartModelFilteredTitle: 'Top Model - MONTH',
    chartModelFilteredSub: 'Cột: Top Model doanh số (K$) theo bộ lọc thời gian',
    chartCustomerStackedTitle: 'ALL YEAR (CUSTOM)',
    chartCustomerStackedSub: 'Cột chồng: Doanh số Khách hàng (K$) theo từng Năm',
    chartCustomerDonutTitle: 'MONTH (CUSTOM)',
    chartCustomerDonutSub: 'Tỷ trọng theo Khách hàng theo bộ lọc thời gian',
    prodLabel: 'Sản xuất',
    shipLabel: 'Xuất hàng',
    salesLabel: 'Doanh số',
    tableTitle: 'Top 15 Model theo Doanh số',
    colModel: 'Model',
    colPartner: 'Custom',
    colCustomer: 'Khách hàng',
    colType: 'TYPE',
    colProd: 'Sản xuất (K)',
    colShip: 'Xuất hàng (K)',
    colSales: 'Doanh số (K$)',
    colRatio: 'Tỷ lệ XH/SX',
    tableEmpty: 'Không có dữ liệu phù hợp với bộ lọc hiện tại.',
    dbTitle: 'Toàn bộ Cơ sở Dữ liệu từ Excel',
    dbSearchPlace: 'Tìm kiếm Model, Đối tác, Khách hàng, Loại...',
    colDivision: 'Phân loại',
    colYear: 'Năm',
    colMonth: 'Tháng',
    colValue: 'Giá trị',
    showingRows: 'Hiển thị {start} đến {end} trong số {total} dòng',
    rowsPerPage: 'Dòng/trang:'
  },
  en: {
    marketingInsights: 'Marketing Insights',
    mainTitleUpload: 'Production – Shipment – Sales Overview',
    mainTitleDash: 'Production – Shipment – Sales Overview',
    subtitleUpload: 'Upload Excel file to analyze by Model, Partner, Customer, and time',
    subtitleDash: 'Analyze by Model, Partner, Customer, and time',
    dragDrop: 'Drag & drop Excel file here',
    orClickBrowse: 'or click to browse from your computer',
    selectExcelBtn: 'Select Excel File (.xlsx)',
    sampleDataBtn: 'Sample Data (Marketing)',
    salesDataBtn: 'Sales Data (Test 2.xlsx)',
    formatsHint: 'Supports .xlsx and .xls formats',
    browserProcessOnly: 'Data is processed directly on your browser, no servers involved.',
    invalidFileError: 'Only .xlsx or .xls files are accepted',
    loadingError: 'Cannot read file. The file may be corrupted or unsupported.',
    fromYear: 'From Year',
    toYear: 'To Year',
    partner: 'Custom',
    customer: 'Customer',
    resetBtn: 'Reset',
    loadExcelBtn: 'Load other Excel',
    filtersApplied: 'filters applied',
    rowsPill: 'rows of data',
    allOption: 'All',
    totalSales: 'Total Sales',
    totalShipment: 'Total Shipment',
    totalProduction: 'Total Production',
    conversionRatio: 'Shipment / Production Ratio',
    salesGrowthYoY: 'Sales Growth (YoY)',
    insufficientYears: 'Not enough data for 2 full years',
    growthComparison: 'vs',
    chart1Title: 'Shipment & Sales by Year',
    chart1Sub: 'Bar: Shipment (K) · Line: Sales (K$) · Dashed: Growth (YoY)',
    chartMonthTitle: 'Shipment & Sales by Month',
    chartMonthSub: 'Bar: Shipment (K) · Line: Sales (K$) · Dashed: Growth (MoM)',
    chart1BarName: 'Shipment (K)',
    chart1LineName: 'Sales (K$)',
    chart2Title: 'Production → Shipment → Sales Funnel',
    chart2Sub: 'Total within selected filters',
    chart3Title: 'Share by Partner',
    chart3Sub: 'By Sales',
    chart4Title: 'Shipment / Production Ratio',
    chart4Sub: 'Conversion efficiency',
    chart5Title: 'Top 10 Models by Sales',
    chart5Sub: 'K$ · sorted descending',
    chart6Title: 'Production by Type (TYPE)',
    chart6Sub: 'Total output (K)',
    chart7Title: 'Share by Customer',
    chart7Sub: 'By Sales',
    chartModelAllYearTitle: 'Top Model - ALL YEAR',
    chartModelAllYearSub: 'Bar: Sales (K$) by Model · Blue: Best Sales · Gold: Other',
    chartModelFilteredTitle: 'Top Model - MONTH',
    chartModelFilteredSub: 'Bar: Sales (K$) by Model for selected period',
    chartCustomerStackedTitle: 'ALL YEAR (CUSTOM)',
    chartCustomerStackedSub: 'Stacked Bar: Customer Sales (K$) by Year',
    chartCustomerDonutTitle: 'MONTH (CUSTOM)',
    chartCustomerDonutSub: 'Share by Customer for selected period',
    prodLabel: 'Production',
    shipLabel: 'Shipment',
    salesLabel: 'Sales',
    tableTitle: 'Top 15 Models by Sales',
    colModel: 'Model',
    colPartner: 'Custom',
    colCustomer: 'Customer',
    colType: 'TYPE',
    colProd: 'Production (K)',
    colShip: 'Shipment (K)',
    colSales: 'Sales (K$)',
    colRatio: 'Ship/Prod Ratio',
    tableEmpty: 'No data matches the current filters.',
    dbTitle: 'Full Database from Excel',
    dbSearchPlace: 'Search Model, Partner, Customer, Type...',
    colDivision: 'Division',
    colYear: 'Year',
    colMonth: 'Month',
    colValue: 'Value',
    showingRows: 'Showing {start} to {end} of {total} rows',
    rowsPerPage: 'Rows/page:'
  },
  ko: {
    marketingInsights: '마케팅 인사이트',
    mainTitleUpload: '생산 – 출하 – 매출 현황',
    mainTitleDash: '생산 – 출하 – 매출 현황',
    subtitleUpload: 'Excel 파일을 업로드하여 모델, 파트너, 고객 및 시간별로 분석합니다',
    subtitleDash: '모델, 파트너, 고객 및 시간별 분석',
    dragDrop: 'Excel 파일을 여기에 드래그 앤 드롭하세요',
    orClickBrowse: '또는 클릭하여 컴퓨터에서 찾아보세요',
    selectExcelBtn: 'Excel 파일 선택 (.xlsx)',
    sampleDataBtn: '샘플 데이터 (마케팅)',
    salesDataBtn: '매출 데이터 (Test 2.xlsx)',
    formatsHint: '.xlsx 및 .xls 형식 지원',
    browserProcessOnly: '데이터는 브라우저에서 직접 처리되며 서버로 전송되지 않습니다.',
    invalidFileError: '.xlsx 또는 .xls 파일만 접수 가능합니다',
    loadingError: '파일을 읽을 수 없습니다. 파일이 손상되었거나 지원되지 않는 형식일 수 있습니다.',
    fromYear: '시작 연도',
    toYear: '종료 연도',
    partner: 'Custom',
    customer: '고객',
    resetBtn: '초기화',
    loadExcelBtn: '다른 Excel 업로드',
    filtersApplied: '개의 필터 적용됨',
    rowsPill: '행의 데이터',
    allOption: '전체',
    totalSales: '총 매출',
    totalShipment: '총 출하',
    totalProduction: '총 생산',
    conversionRatio: '생산 대비 출하율',
    salesGrowthYoY: '매출 성장률 (YoY)',
    insufficientYears: '2년 미만의 데이터',
    growthComparison: '대비',
    chart1Title: '연도별 출하 및 매출 추이',
    chart1Sub: '막대: 출하 (K) · 선: 매출 (K$) · 점선: 전년 대비 증감',
    chartMonthTitle: '월별 출하 및 매출 추이',
    chartMonthSub: '막대: 출하 (K) · 선: 매출 (K$) · 점선: 전월 대비 증감',
    chart1BarName: '출하 (K)',
    chart1LineName: '매출 (K$)',
    chart2Title: '생산 → 출하 → 매출 깔대기',
    chart2Sub: '선택한 필터 범위의 총계',
    chart3Title: '파트너별 비중',
    chart3Sub: '매출액 기준',
    chart4Title: '생산 대비 출하율',
    chart4Sub: '전환 효율',
    chart5Title: '매출 기준 Top 10 모델',
    chart5Sub: 'K$ · 내림차순 정렬',
    chart6Title: '유형별 생산 (TYPE)',
    chart6Sub: '총 생산량 (K)',
    chart7Title: '고객별 비중',
    chart7Sub: '매출액 기준',
    chartModelAllYearTitle: 'Top Model - ALL YEAR',
    chartModelAllYearSub: '막대: 모델별 매출 (K$) · 파란색: 베스트 셀러 · 노란색: 기타',
    chartModelFilteredTitle: 'Top Model - MONTH',
    chartModelFilteredSub: '막대: 선택 기간 모델별 매출 (K$)',
    chartCustomerStackedTitle: 'ALL YEAR (CUSTOM)',
    chartCustomerStackedSub: '누적 막대: 연도별 고객사 매출 추이 (K$)',
    chartCustomerDonutTitle: 'MONTH (CUSTOM)',
    chartCustomerDonutSub: '선택 기간 고객사별 매출 비중',
    prodLabel: '생산',
    shipLabel: '출하',
    salesLabel: '매출',
    tableTitle: '매출 기준 Top 15 모델',
    colModel: '모델',
    colPartner: 'Custom',
    colCustomer: '고객',
    colType: '유형',
    colProd: '생산 (K)',
    colShip: '출하 (K)',
    colSales: '매출 (K$)',
    colRatio: '출하/생산 비율',
    tableEmpty: '현재 필터와 일치하는 데이터가 없습니다.',
    dbTitle: 'Excel 전체 데이터베이스',
    dbSearchPlace: '모델, 파트너, 고객, 유형 검색...',
    colDivision: '구분',
    colYear: '연도',
    colMonth: '월',
    colValue: '값',
    showingRows: '{total}개 중 {start}~{end} 행 표시',
    rowsPerPage: '페이지당 행 수:'
  }
};
