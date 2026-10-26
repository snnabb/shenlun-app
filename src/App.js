import React, { useRef, useState } from "react";
import {
  AlignmentType,
  BorderStyle,
  Document,
  HeightRule,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  TableLayoutType,
  VerticalAlign,
  WidthType,
} from "docx";
import html2canvas from "html2canvas";
import { jsPDF } from "jspdf";
import "./App.css";

const COLS = 25;
const DRAFT_STORAGE_KEY = "shenlun-app-draft-v1";

const createEmptyGrid = (rowCount) =>
  Array.from({ length: rowCount }, () => Array(COLS).fill(""));

const getCellMarker = (rowIndex, colIndex) => {
  const charCount = rowIndex * COLS + colIndex + 1;
  return charCount % 100 === 0 ? charCount : null;
};

const getDefaultFilename = () => {
  const date = new Date().toLocaleDateString("zh-CN").replace(/\//g, "-");
  return `申论答题-${date}`;
};

const sanitizeFilename = (input, extension) => {
  const safeName = input.trim().replace(/[\\/:*?"<>|]/g, "_");
  return `${safeName || getDefaultFilename()}.${extension}`;
};

const downloadBlob = (blob, filename) => {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
};

const loadDraft = () => {
  try {
    const rawDraft = window.localStorage.getItem(DRAFT_STORAGE_KEY);
    if (!rawDraft) return null;

    const draft = JSON.parse(rawDraft);
    const savedRows = Number.parseInt(draft.rows, 10);
    if (!Number.isInteger(savedRows) || savedRows < 1 || !Array.isArray(draft.grid)) {
      return null;
    }

    const grid = Array.from({ length: savedRows }, (_, rowIndex) =>
      Array.from({ length: COLS }, (_, colIndex) => {
        const value = draft.grid[rowIndex]?.[colIndex];
        return typeof value === "string" ? value : "";
      })
    );

    return { rows: savedRows, grid };
  } catch {
    return null;
  }
};

const hasSavedDraft = () => loadDraft() !== null;

const App = () => {
  const cols = COLS;
  const [rows, setRows] = useState(8);
  const [grid, setGrid] = useState(() => createEmptyGrid(8));
  const [draftAvailable, setDraftAvailable] = useState(hasSavedDraft);
  const [saveMessage, setSaveMessage] = useState("");
  const [filenameDialog, setFilenameDialog] = useState(null);
  const [filenameInput, setFilenameInput] = useState("");

  // 中文输入法可能一次提交多个字符。用元素引用跟踪组合输入，避免
  // 在 composition 期间提前提交。
  const composingCellRef = useRef(null);
  const committedCompositionRef = useRef(null);

  const isValidInput = (value) =>
    value.length > 0 && !/[\r\n]/u.test(value);

  const focusCell = (row, col) => {
    const nextCell = document.querySelector(
      `[data-row="${row}"][data-col="${col}"]`
    );
    if (nextCell) {
      nextCell.focus();
    }
  };

  const updateGrid = (row, col, value) => {
    setGrid((currentGrid) => {
      const updatedGrid = currentGrid.map((gridRow) => [...gridRow]);
      updatedGrid[row][col] = value;
      return updatedGrid;
    });
  };

  const commitText = (value, row, col) => {
    if (value === "") {
      updateGrid(row, col, "");
      return;
    }

    if (!isValidInput(value)) {
      updateGrid(row, col, "");
      return;
    }

    // 一个格子只保存一个字符；一次提交多个字符时依次填入后续格子。
    const chars = Array.from(value);
    const updatedGrid = grid.map((gridRow) => [...gridRow]);
    let currentRow = row;
    let currentCol = col;

    for (const char of chars) {
      if (currentRow >= rows) break;

      updatedGrid[currentRow][currentCol] = char;
      currentCol += 1;

      if (currentCol >= cols) {
        currentCol = 0;
        currentRow += 1;
      }
    }

    setGrid(updatedGrid);
    focusCell(currentRow, currentCol);
  };

  const handleInputChange = (e, row, col) => {
    const value = e.target.value;

    if (
      composingCellRef.current === e.currentTarget ||
      e.nativeEvent.isComposing
    ) {
      // 这是受控 input：组合期间也要保留临时值，否则 React 会把输入框
      // 恢复成旧的空值，导致 compositionend 拿不到最终字符。
      updateGrid(row, col, value);
      return;
    }

    // compositionend 后部分浏览器还会补发一次 onChange；组合结果已经
    // 提交过时跳过这一次，避免空值把刚提交的文字清掉。
    if (committedCompositionRef.current?.element === e.currentTarget) {
      committedCompositionRef.current = null;
      return;
    }

    committedCompositionRef.current = null;
    commitText(value, row, col);
  };

  const handleCompositionStart = (e) => {
    composingCellRef.current = e.currentTarget;
  };

  const handleCompositionEnd = (e, row, col) => {
    composingCellRef.current = null;

    const value = e.currentTarget.value;
    const data = e.nativeEvent.data || "";
    const compositionText =
      Array.from(data).length >= Array.from(value).length ? data : value;

    if (compositionText !== "") {
      committedCompositionRef.current = {
        element: e.currentTarget,
      };
      commitText(compositionText, row, col);
    }
  };

  const moveToNextCell = (row, col) => {
    const nextCol = col + 1;
    const nextRow = nextCol >= cols ? row + 1 : row;
    const actualCol = nextCol >= cols ? 0 : nextCol;

    const nextCell = document.querySelector(
      `[data-row="${nextRow}"][data-col="${actualCol}"]`
    );
    if (nextCell) {
      nextCell.focus();
    }
  };

  const handleKeyDown = (e, row, col) => {
    if (e.key === "Backspace") {
      if (grid[row][col] === "") {
        moveToPreviousCell(row, col);
      } else {
        updateGrid(row, col, "");
      }
    }
  };

  const moveToPreviousCell = (row, col) => {
    const prevCol = col - 1;
    const prevRow = prevCol < 0 ? row - 1 : row;
    const actualCol = prevCol < 0 ? cols - 1 : prevCol;

    const prevCell = document.querySelector(
      `[data-row="${prevRow}"][data-col="${actualCol}"]`
    );
    if (prevCell) {
      prevCell.focus();
    }
  };

  const handlePaste = (e, row, col) => {
    e.preventDefault();
    const pasteText = e.clipboardData.getData("text");
    const cells = pasteText.split("");

    let currentRow = row;
    let currentCol = col;

    const updatedGrid = grid.map((gridRow) => [...gridRow]);

    for (let i = 0; i < cells.length; i++) {
      if (currentRow >= rows) break;
      if (currentCol >= cols) {
        currentCol = 0;
        currentRow++;
      }

      updatedGrid[currentRow][currentCol] = cells[i];
      currentCol++;
    }

    setGrid(updatedGrid);

    moveToNextCell(currentRow, currentCol - 1);
  };

  const handleRowChange = (e) => {
    const newRowCount = parseInt(e.target.value, 10);
    const updatedGrid = Array.from({ length: newRowCount }, (_, rowIndex) =>
      grid[rowIndex] ? grid[rowIndex] : Array(cols).fill("")
    );
    setRows(newRowCount);
    setGrid(updatedGrid);
  };

  const clearGrid = () => {
    setGrid(createEmptyGrid(rows));
  };

  const getGridText = () =>
    grid.map((row) => row.filter((cell) => cell !== "").join("")).join("\n");

  const saveDraft = () => {
    try {
      window.localStorage.setItem(
        DRAFT_STORAGE_KEY,
        JSON.stringify({ rows, grid, savedAt: new Date().toISOString() })
      );
      setDraftAvailable(true);
      setSaveMessage("草稿已保存到当前浏览器");
    } catch {
      setSaveMessage("保存失败，请改用下载文本");
    }
  };

  const restoreDraft = () => {
    const savedDraft = loadDraft();
    if (!savedDraft) {
      setDraftAvailable(false);
      setSaveMessage("当前浏览器没有可恢复的草稿");
      return;
    }

    setRows(savedDraft.rows);
    setGrid(savedDraft.grid);
    setDraftAvailable(true);
    setSaveMessage("已恢复上次保存的草稿");
  };

  const exportWord = async (filename) => {
    try {
      const border = {
        style: BorderStyle.SINGLE,
        size: 4,
        color: "FF0000",
      };
      const tableRows = grid.map(
        (row, rowIndex) =>
          new TableRow({
            height: { value: 400, rule: HeightRule.EXACTLY },
            children: row.map((cell, colIndex) => {
              const marker = getCellMarker(rowIndex, colIndex);
              const children = [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { before: 0, after: 0, line: 180 },
                  children: [
                    new TextRun({
                      text: cell || " ",
                      font: "Microsoft YaHei",
                      size: 18,
                    }),
                  ],
                }),
              ];

              if (marker) {
                children.push(
                  new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    spacing: { before: 0, after: 0, line: 80 },
                    children: [
                      new TextRun({
                        text: String(marker),
                        font: "Arial",
                        size: 8,
                        color: "666666",
                      }),
                    ],
                  })
                );
              }

              return new TableCell({
                width: { size: 400, type: WidthType.DXA },
                margins: { top: 0, bottom: 0, left: 0, right: 0 },
                verticalAlign: VerticalAlign.CENTER,
                borders: {
                  top: border,
                  bottom: border,
                  left: border,
                  right: border,
                },
                children,
              });
            }),
          })
      );

      const documentFile = new Document({
        sections: [
          {
            properties: {
              page: {
                size: { width: 11906, height: 16838 },
                margin: { top: 720, right: 720, bottom: 720, left: 720 },
              },
            },
            children: [
              new Table({
                width: { size: 10000, type: WidthType.DXA },
                layout: TableLayoutType.FIXED,
                columnWidths: Array(cols).fill(400),
                rows: tableRows,
              }),
            ],
          },
        ],
      });

      const blob = await Packer.toBlob(documentFile);
      downloadBlob(blob, filename);
      setSaveMessage("Word 文件已下载");
    } catch {
      setSaveMessage("Word 下载失败，请改用 PDF 或下载文本");
    }
  };

  const createPdfExportElement = () => {
    const exportGrid = document.createElement("div");
    exportGrid.className = "pdf-export-grid";

    grid.forEach((row, rowIndex) => {
      const exportRow = document.createElement("div");
      exportRow.className = "pdf-export-row";
      row.forEach((cell, colIndex) => {
        const wrapper = document.createElement("div");
        wrapper.className = "pdf-export-cell";
        const text = document.createElement("span");
        text.textContent = cell || "";
        wrapper.appendChild(text);
        const marker = getCellMarker(rowIndex, colIndex);
        if (marker) {
          const markerElement = document.createElement("small");
          markerElement.textContent = String(marker);
          wrapper.appendChild(markerElement);
        }
        exportRow.appendChild(wrapper);
      });
      exportGrid.appendChild(exportRow);
    });

    document.body.appendChild(exportGrid);
    return exportGrid;
  };

  const exportPdf = async (filename) => {
    let exportGrid;
    try {
      setSaveMessage("正在生成 PDF…");
      exportGrid = createPdfExportElement();
      const canvas = await html2canvas(exportGrid, {
        backgroundColor: "#ffffff",
        scale: 2,
        useCORS: true,
        logging: false,
      });
      const pdf = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
      const pageMargin = 10;
      const pageWidth = 210 - pageMargin * 2;
      const pageHeight = 297 - pageMargin * 2;
      const pageHeightPx = Math.floor((canvas.width * pageHeight) / pageWidth);

      for (let offset = 0, pageIndex = 0; offset < canvas.height; offset += pageHeightPx, pageIndex += 1) {
        const sliceHeight = Math.min(pageHeightPx, canvas.height - offset);
        const pageCanvas = document.createElement("canvas");
        pageCanvas.width = canvas.width;
        pageCanvas.height = sliceHeight;
        pageCanvas.getContext("2d").drawImage(
          canvas,
          0,
          offset,
          canvas.width,
          sliceHeight,
          0,
          0,
          canvas.width,
          sliceHeight
        );
        if (pageIndex > 0) pdf.addPage();
        pdf.addImage(
          pageCanvas.toDataURL("image/png"),
          "PNG",
          pageMargin,
          pageMargin,
          pageWidth,
          (sliceHeight * pageWidth) / canvas.width
        );
      }

      pdf.save(filename);
      setSaveMessage("PDF 文件已直接下载");
    } catch {
      setSaveMessage("PDF 下载失败，请重试");
    } finally {
      exportGrid?.remove();
    }
  };

  const openFilenameDialog = (type, extension) => {
    setFilenameInput(getDefaultFilename());
    setFilenameDialog({ type, extension });
  };

  const cancelFilenameDialog = () => {
    setFilenameDialog(null);
  };

  const submitFilenameDialog = () => {
    if (!filenameDialog) return;

    const { type, extension } = filenameDialog;
    const filename = sanitizeFilename(filenameInput, extension);
    setFilenameDialog(null);

    if (type === "word") {
      exportWord(filename);
    } else {
      void exportPdf(filename);
    }
  };

  const copyGrid = () => {
    const text = grid.flat().join("");
    if (text === "" || text === undefined || text == null) {
      return;
    }
    navigator.clipboard.writeText(getGridText());
  };

  return (
    <div className="app-container">
      <div className="controls">
        <label htmlFor="row-selector">选择字数：</label>
        <select id="row-selector" onChange={handleRowChange} value={rows}>
          <option value="8">200字</option>
          <option value="10">250字</option>
          <option value="12">300字</option>
          <option value="14">350字</option>
          <option value="16">400字</option>
          <option value="18">450字</option>
          <option value="20">500字</option>
          <option value="48">1200字</option>
        </select>
      </div>
      <div className="grid" id="grid-container">
        {grid.map((row, rowIndex) => (
          <div className="grid-row" key={rowIndex}>
            {row.map((cell, colIndex) => (
              <div
                className="grid-cell-wrapper"
                key={`${rowIndex}-${colIndex}`}
              >
                <input
                  data-row={rowIndex}
                  data-col={colIndex}
                  className="grid-cell"
                  value={cell}
                  onChange={(e) => handleInputChange(e, rowIndex, colIndex)}
                  onCompositionStart={handleCompositionStart}
                  onCompositionEnd={(e) =>
                    handleCompositionEnd(e, rowIndex, colIndex)
                  }
                  onKeyDown={(e) => handleKeyDown(e, rowIndex, colIndex)}
                  onPaste={(e) => handlePaste(e, rowIndex, colIndex)}
                />
                {getCellMarker(rowIndex, colIndex) && (
                  <span className="char-count-marker" aria-hidden="true">
                    {getCellMarker(rowIndex, colIndex)}
                  </span>
                )}
              </div>
            ))}
          </div>
        ))}
      </div>
      <div className="buttons">
        <button onClick={saveDraft}>保存草稿</button>
        <button onClick={restoreDraft} disabled={!draftAvailable}>
          恢复草稿
        </button>
        <button onClick={() => openFilenameDialog("word", "docx")}>
          下载 Word
        </button>
        <button onClick={() => openFilenameDialog("pdf", "pdf")}>
          下载 PDF
        </button>
        <button onClick={copyGrid}>复制内容</button>
        <button onClick={clearGrid}>清空内容</button>
      </div>
      {saveMessage && <div className="save-message">{saveMessage}</div>}
      {filenameDialog && (
        <div className="filename-dialog-backdrop">
          <div
            className="filename-dialog"
            role="dialog"
            aria-modal="true"
            aria-labelledby="filename-dialog-title"
          >
            <h2 id="filename-dialog-title">输入文件名</h2>
            <div className="filename-input-row">
              <input
                autoFocus
                value={filenameInput}
                onChange={(e) => setFilenameInput(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === "Enter") submitFilenameDialog();
                  if (e.key === "Escape") cancelFilenameDialog();
                }}
              />
              <span>.{filenameDialog.extension}</span>
            </div>
            <div className="filename-dialog-buttons">
              <button onClick={cancelFilenameDialog}>取消</button>
              <button onClick={submitFilenameDialog}>确定下载</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default App;
