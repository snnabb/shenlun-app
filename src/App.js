import React, { useRef, useState } from "react";
import "./App.css";

const COLS = 25;
const DRAFT_STORAGE_KEY = "shenlun-app-draft-v1";

const createEmptyGrid = (rowCount) =>
  Array.from({ length: rowCount }, () => Array(COLS).fill(""));

const loadDraft = () => {
  const fallback = { rows: 8, grid: createEmptyGrid(8), restored: false };

  try {
    const rawDraft = window.localStorage.getItem(DRAFT_STORAGE_KEY);
    if (!rawDraft) return fallback;

    const draft = JSON.parse(rawDraft);
    const savedRows = Number.parseInt(draft.rows, 10);
    if (!Number.isInteger(savedRows) || savedRows < 1 || !Array.isArray(draft.grid)) {
      return fallback;
    }

    const grid = Array.from({ length: savedRows }, (_, rowIndex) =>
      Array.from({ length: COLS }, (_, colIndex) => {
        const value = draft.grid[rowIndex]?.[colIndex];
        return typeof value === "string" ? value : "";
      })
    );

    return { rows: savedRows, grid, restored: true };
  } catch {
    return fallback;
  }
};

const App = () => {
  const [initialDraft] = useState(loadDraft);
  const cols = COLS;
  const [rows, setRows] = useState(initialDraft.rows);
  const [grid, setGrid] = useState(initialDraft.grid);
  const [saveMessage, setSaveMessage] = useState(
    initialDraft.restored ? "已恢复上次保存的草稿" : ""
  );

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
      setSaveMessage("草稿已保存到当前浏览器");
    } catch {
      setSaveMessage("保存失败，请改用下载文本");
    }
  };

  const downloadText = () => {
    const text = getGridText();
    if (!text.replace(/\n/g, "")) {
      setSaveMessage("当前没有可保存的内容");
      return;
    }

    const blob = new Blob([`${text}\n`], {
      type: "text/plain;charset=utf-8",
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    const date = new Date().toLocaleDateString("zh-CN").replace(/\//g, "-");
    link.href = url;
    link.download = `申论答题-${date}.txt`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
    setSaveMessage("文本已下载");
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
              <input
                key={`${rowIndex}-${colIndex}`}
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
            ))}
          </div>
        ))}
      </div>
      <div className="buttons">
        <button onClick={saveDraft}>保存草稿</button>
        <button onClick={downloadText}>下载文本</button>
        <button onClick={copyGrid}>复制内容</button>
        <button onClick={clearGrid}>清空内容</button>
      </div>
      {saveMessage && <div className="save-message">{saveMessage}</div>}
    </div>
  );
};

export default App;
