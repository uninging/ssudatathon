const content = document.getElementById("content");
let is_btn1 = false;

function renderInitialScreen() {
  content.innerHTML = `
    <div class="button-container">
      <button class="button" id="button1">보존서고 이동 도서<br>탐색하기</button>
      <button class="button" id="button2">서가로 복귀시킬 도서<br>탐색하기</button>
    </div>
    <h1>원하시는 작업을 선택하세요</h1>
  `;

  document.getElementById("button1").addEventListener("click", () => {
    is_btn1 = true;
    renderFileSelectionScreen("보존서고로 이동시킬 책을 탐색합니다");
  });

  document.getElementById("button2").addEventListener("click", () => {
    is_btn1 = false;
    renderFileSelectionScreen("서가로 복귀시킬 책을 탐색합니다");
  });
}

function renderFileSelectionScreen(title) {
  content.innerHTML = `
    <h1>${title}</h1>
    <div class="file-actions">
      <div class="drop-zone" id="dropZone1">
        단행본(도서)정보.xlsx 파일을<br>드래그 앤 드롭하거나<br>클릭하여 선택하세요.
      </div>
      <input type="file" id="fileInput1" hidden />
      <div class="drop-zone" id="dropZone2">
        대출정보.xlsx 파일을<br>드래그 앤 드롭하거나<br>클릭하여 선택하세요.
      </div>
      <input type="file" id="fileInput2" hidden />
    </div>
    <button id="selectButton" class="button" style="display:none;">선정하기</button>
    <button class="back-button" id="backButton">뒤로가기</button>
  `;

  document.getElementById("backButton").addEventListener("click", renderInitialScreen);
  setupDragAndDrop("dropZone1", "fileInput1", "dropZone2", "fileInput2");
}

function setupDragAndDrop(dropZone1Id, fileInput1Id, dropZone2Id, fileInput2Id) {
  const dropZone1 = document.getElementById(dropZone1Id);
  const fileInput1 = document.getElementById(fileInput1Id);
  const dropZone2 = document.getElementById(dropZone2Id);
  const fileInput2 = document.getElementById(fileInput2Id);
  const selectButton = document.getElementById("selectButton");

  let file1 = null, file2 = null;

  function handleFileSelection(fileInput, setFile, dropZone) {
    fileInput.addEventListener("change", () => {
      setFile(fileInput.files[0]);
      dropZone.textContent = `${fileInput.files[0].name} 선택됨.`;
      toggleSelectButton();
    });

    dropZone.addEventListener("click", () => fileInput.click());

    dropZone.addEventListener("dragover", (e) => {
      e.preventDefault();
      dropZone.classList.add("dragover");
    });

    dropZone.addEventListener("dragleave", () => {
      dropZone.classList.remove("dragover");
    });

    dropZone.addEventListener("drop", (e) => {
      e.preventDefault();
      dropZone.classList.remove("dragover");
      const file = e.dataTransfer.files[0];
      setFile(file);
      dropZone.textContent = `${file.name} 선택됨.`;
      toggleSelectButton();
    });
  }

  handleFileSelection(fileInput1, (file) => file1 = file, dropZone1);
  handleFileSelection(fileInput2, (file) => file2 = file, dropZone2);

  function toggleSelectButton() {
    const selectButton = document.getElementById("selectButton");
    const backButton = document.getElementById("backButton");
    
    if (file1 && file2) {
      selectButton.style.display = "inline-block";
      backButton.style.display = "none"; // 뒤로가기 숨김
    } else {
      selectButton.style.display = "none";
      backButton.style.display = "inline-block"; // 뒤로가기 표시
    }
  }
  
  selectButton.addEventListener("click", () => {
    if (file1 && file2) {
      content.innerHTML = `
        <h1>생성 중입니다...</h1>
        <p>잠시만 기다려 주세요.</p>
        <img src="https://github.com/uninging/ssudatathon/blob/main/Fading_circles.gif?raw=true" alt="로딩 중" />
      `;
      processFilesWrapper(file1, file2);
    }
  });
}

async function processFilesWrapper(file1, file2) {
  try {
    if (is_btn1) {
      await processFiles(file1, file2); // 보존서고로 이동 조건 실행
    } else {
      await processFiles1(file1, file2); // 서가로 복귀 조건 실행
    }
    
    // 파일 생성 완료 후 표시
    content.innerHTML = `
      <h1>파일 생성이 완료되었습니다!</h1>
      <p>파일이 다운로드되었습니다.</p>
      <button class="back-button" id="backToStartButton">처음으로 돌아가기</button>
    `;
    document.getElementById("backToStartButton").addEventListener("click", renderInitialScreen);
  } catch (error) {
    // 오류 발생 시 표시
    content.innerHTML = `
      <h1>오류 발생</h1>
      <p>${error.message}</p>
      <button class="back-button" id="backToStartButton">처음으로 돌아가기</button>
    `;
    document.getElementById("backToStartButton").addEventListener("click", renderInitialScreen);
  }
}

async function processFilesWrapper(file1, file2) {
  content.innerHTML = `
    <h1>생성 중입니다...</h1>
    <p>잠시만 기다려 주세요.</p>
    <img src="https://raw.githubusercontent.com/uninging/ssudatathon/main/Fading_circles.gif" alt="로딩 중" style="display: block; margin: 0 auto;"/>
  `;
  try {
    if (is_btn1) {
      await processFiles(file1, file2); // 보존서고로 이동 조건 실행
    } else {
      await processFiles1(file1, file2); // 서가로 복귀 조건 실행
    }
    
    // 파일 생성 완료 후 표시
    content.innerHTML = `
      <h1>파일 생성이 완료되었습니다!</h1>
      <p>파일이 다운로드되었습니다.</p>
      <button class="back-button" id="backToStartButton">처음으로 돌아가기</button>
    `;
    document.getElementById("backToStartButton").addEventListener("click", renderInitialScreen);
  } catch (error) {
    // 오류 발생 시 표시
    content.innerHTML = `
      <h1>오류 발생</h1>
      <p>${error.message}</p>
      <button class="back-button" id="backToStartButton">처음으로 돌아가기</button>
    `;
    document.getElementById("backToStartButton").addEventListener("click", renderInitialScreen);
  }
}


async function processFiles(file1, file2) {
  const [data1, data2] = await Promise.all([readExcel(file1), readExcel(file2)]);
  // 보존서고 이동 로직
  const books = data1.filter(row => row["소장위치"] === "4층인문");
  const threeYearsAgo = new Date();
  threeYearsAgo.setFullYear(threeYearsAgo.getFullYear() - 3);
  const threeYearsAgoExcel = convertDateToExcelDate(threeYearsAgo);
  const loans = data2.filter(row => new Date(row["대출일시"]) >= threeYearsAgoExcel);

  const loanCounts = loans.reduce((acc, row) => {
    acc[row["도서ID"]] = (acc[row["도서ID"]] || 0) + 1;
    return acc;
}, {});

// 각 도서의 대출 횟수 추가
books.forEach(book => {
    book["대출횟수"] = loanCounts[book["도서ID"]] || 0;
});

// 도서 그룹화 (ISBN이 없을 경우 서명을 기준으로 그룹화)
const groupedBooks = Object.values(books.reduce((acc, book) => {
    const key = book["ISBN"] && book["ISBN"] !== "0" ? book["ISBN"] : book["서명"];

    if (!acc[key]) {
        acc[key] = { ...book, "대출횟수": 0 };
    }
    acc[key]["대출횟수"] += book["대출횟수"];

    return acc;
}, {}));

console.log(groupedBooks);


  const filteredBooks = groupedBooks.filter(book => {
    const registrationDateExcel = book["등록일자"];
    const fiveYearsAgo = new Date();
    fiveYearsAgo.setFullYear(fiveYearsAgo.getFullYear() - 5);
    const fiveYearsAgoExcel = convertDateToExcelDate(fiveYearsAgo);
    return registrationDateExcel <= fiveYearsAgoExcel && book["대출횟수"] <= 2;
  });

  createExcelFile(filteredBooks, "GoToClosedRepository.xlsx");
}

async function processFiles1(file1, file2) {
  const [data1, data2] = await Promise.all([readExcel(file1), readExcel(file2)]);
  // 서가 복귀 로직
  const books = data1.filter(row => row["소장위치"] === "보존서고");
  const threeYearsAgo = new Date();
  threeYearsAgo.setFullYear(threeYearsAgo.getFullYear() - 3);
  const threeYearsAgoExcel = convertDateToExcelDate(threeYearsAgo);
  const loans = data2.filter(row => new Date(row["대출일시"]) >= threeYearsAgoExcel);

  const loanCounts = loans.reduce((acc, row) => {
    acc[row["도서ID"]] = (acc[row["도서ID"]] || 0) + 1;
    return acc;
  }, {});

books.forEach(book => {
    book["대출횟수"] = loanCounts[book["도서ID"]] || 0;
});

const groupedBooks = Object.values(books.reduce((acc, book) => {
    const isbn = book["ISBN"];
    const key = (isbn && isbn !== "0") ? isbn : book["서명"];  // ISBN이 없으면 서명을 키로 사용

    if (!acc[key]) {
        acc[key] = { ...book, "대출횟수": 0 };
    }
    acc[key]["대출횟수"] += book["대출횟수"];
    return acc;
}, {}));

console.log(groupedBooks);


  const filteredBooks = groupedBooks.filter(book => book["대출횟수"] >= 24);

  createExcelFile(filteredBooks, "GoToBookshelf.xlsx");
}

function createExcelFile(data, fileName) {
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  XLSX.writeFile(wb, fileName);
}

function readExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const data = reader.result;
      const workbook = XLSX.read(data, { type: "binary" });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(sheet);
      resolve(jsonData);
    };
    reader.onerror = (error) => reject(error);
    reader.readAsBinaryString(file);
  });
}

function convertDateToExcelDate(date) {
  return Math.floor((date - new Date(1899, 11, 30)) / (1000 * 60 * 60 * 24));
}

// 초기 화면 렌더링
renderInitialScreen();
