import React, { useState, useCallback, useRef, useEffect } from "react";
import { FixedSizeGrid as Grid } from "react-window";
import "./Excel.css";

const GRID_SIZE = 10000;
const CELL_SIZE = 100; // Width of regular cells
const HEADER_SIZE = 40; // Height of header cells and width of row headers

const getColumnLabel = (index) => {
  let label = "";
  while (index >= 0) {
    label = String.fromCharCode(65 + (index % 26)) + label;
    index = Math.floor(index / 26) - 1;
  }
  return label;
};

const Excel = () => {
  const [cells, setCells] = useState({});
  const [editingCell, setEditingCell] = useState(null);
  const [focusedCell, setFocusedCell] = useState({ row: 0, col: 0 });
  const [dimensions, setDimensions] = useState({
    width: window.innerWidth - 20,
    height: window.innerHeight - 20,
  });

  const mainGridRef = useRef();
  const columnHeaderRef = useRef();
  const rowHeaderRef = useRef();
  const containerRef = useRef();
  const mainGridOuterRef = useRef();

  useEffect(() => {
    const handleResize = () => {
      if (containerRef.current) {
        const rect = containerRef.current.getBoundingClientRect();
        setDimensions({
          width: rect.width,
          height: rect.height,
        });
      }
    };

    handleResize();
    window.addEventListener("resize", handleResize);
    return () => window.removeEventListener("resize", handleResize);
  }, []);

  const handleScroll = useCallback((scrollInfo) => {
    const { scrollLeft, scrollTop } = scrollInfo;
    
    // Sync column headers
    if (columnHeaderRef.current) {
      const columnGrid = columnHeaderRef.current;
      columnGrid._outerRef.scrollLeft = scrollLeft;
    }

    // Sync row headers
    if (rowHeaderRef.current) {
      const rowGrid = rowHeaderRef.current;
      rowGrid._outerRef.scrollTop = scrollTop;
    }
  }, []);

  const getCellValue = useCallback(
    (rowIndex, columnIndex) => {
      const cellKey = `${rowIndex},${columnIndex}`;
      return cells[cellKey] || "";
    },
    [cells]
  );

  const updateCell = useCallback((rowIndex, columnIndex, value) => {
    setCells((prev) => {
      const newCells = { ...prev };
      const cellKey = `${rowIndex},${columnIndex}`;
      newCells[cellKey] = value;
      return newCells;
    });
  }, []);

  const evaluateFormula = useCallback(
    (formula) => {
      try {
        const evaluatedFormula = formula.replace(/[A-Z]+\d+/g, (match) => {
          const col = match.match(/[A-Z]+/)[0];
          const row = parseInt(match.match(/\d+/)[0]) - 1;
          const colIndex = col
            .split("")
            .reduce((acc, char) => acc * 26 + char.charCodeAt(0) - 65, 0);
          return getCellValue(row, colIndex) || 0;
        });
        return eval(evaluatedFormula);
      } catch (error) {
        return "#ERROR";
      }
    },
    [getCellValue]
  );

  const moveFocus = useCallback((rowDelta, colDelta) => {
    setFocusedCell(prev => {
      const newRow = Math.max(0, Math.min(GRID_SIZE - 1, prev.row + rowDelta));
      const newCol = Math.max(0, Math.min(GRID_SIZE - 1, prev.col + colDelta));
      
      // Scroll the cell into view if needed
      if (mainGridRef.current) {
        mainGridRef.current.scrollToItem({
          columnIndex: newCol,
          rowIndex: newRow,
          align: "center"
        });
      }
      
      return { row: newRow, col: newCol };
    });
  }, []);

  const handleKeyDown = useCallback((e) => {
    if (editingCell) return; // Don't handle navigation while editing

    switch (e.key) {
      case 'Tab':
        e.preventDefault();
        moveFocus(0, e.shiftKey ? -1 : 1);
        break;
      case 'Enter':
        e.preventDefault();
        moveFocus(1, 0);
        break;
      case 'ArrowRight':
        e.preventDefault();
        moveFocus(0, 1);
        break;
      case 'ArrowLeft':
        e.preventDefault();
        moveFocus(0, -1);
        break;
      case 'ArrowUp':
        e.preventDefault();
        moveFocus(-1, 0);
        break;
      case 'ArrowDown':
        e.preventDefault();
        moveFocus(1, 0);
        break;
      default:
        break;
    }
  }, [moveFocus, editingCell]);

  useEffect(() => {
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [handleKeyDown]);

  const Cell = ({ columnIndex, rowIndex, style }) => {
    const cellKey = `${rowIndex},${columnIndex}`;
    const isEditing = editingCell === cellKey;
    const isFocused = focusedCell.row === rowIndex && focusedCell.col === columnIndex;
    const value = getCellValue(rowIndex, columnIndex);
    const displayValue = value.startsWith("=")
      ? evaluateFormula(value.slice(1))
      : value;

    return (
      <div
        className={`cell ${isEditing ? "editing" : ""} ${isFocused ? "focused" : ""}`}
        style={{
          ...style,
          padding: "4px 6px",
          background: isEditing ? "#fff" : "transparent",
          height: HEADER_SIZE,
          display: "flex",
          alignItems: "center",
        }}
        onClick={() => {
          setFocusedCell({ row: rowIndex, col: columnIndex });
          setEditingCell(cellKey);
        }}
        onBlur={() => setEditingCell(null)}
      >
        {isEditing ? (
          <input
            type="text"
            value={value}
            onChange={(e) => updateCell(rowIndex, columnIndex, e.target.value)}
            autoFocus
            className="cell-input"
            onKeyDown={(e) => {
              if (e.key === 'Enter' || e.key === 'Tab') {
                e.stopPropagation(); // Stop event from bubbling up
                e.preventDefault();
                const nextRow = e.key === 'Enter' ? rowIndex + 1 : rowIndex;
                const nextCol = e.key === 'Tab' ? columnIndex + (e.shiftKey ? -1 : 1) : columnIndex;
                setEditingCell(null);
                setFocusedCell({ row: nextRow, col: nextCol });
                
                // Scroll into view if needed
                if (mainGridRef.current) {
                  mainGridRef.current.scrollToItem({
                    columnIndex: nextCol,
                    rowIndex: nextRow,
                    align: "center"
                  });
                }
              }
            }}
          />
        ) : (
          displayValue
        )}
      </div>
    );
  };

  const ColumnHeader = ({ columnIndex, style }) => (
    <div
      className="header-cell column-header"
      style={{
        ...style,
        height: HEADER_SIZE,
        padding: "4px 6px",
      }}
    >
      {getColumnLabel(columnIndex)}
    </div>
  );

  const RowHeader = ({ rowIndex, style }) => (
    <div
      className="header-cell row-header"
      style={{
        ...style,
        height: HEADER_SIZE,
        padding: "4px 6px",
      }}
    >
      {rowIndex + 1}
    </div>
  );

  return (
    <div className="excel-container" ref={containerRef}>
      <div className="excel-grid">
        <div
          className="corner-header"
          style={{ width: HEADER_SIZE, height: HEADER_SIZE }}
        ></div>
        <div className="column-headers">
          <Grid
            ref={columnHeaderRef}
            className="grid header-grid"
            columnCount={GRID_SIZE}
            columnWidth={CELL_SIZE}
            height={HEADER_SIZE}
            rowCount={1}
            rowHeight={HEADER_SIZE}
            width={dimensions.width - HEADER_SIZE}
            style={{ overflow: "hidden" }}
          >
            {ColumnHeader}
          </Grid>
        </div>
        <div className="row-headers">
          <Grid
            ref={rowHeaderRef}
            className="grid header-grid"
            columnCount={1}
            columnWidth={HEADER_SIZE}
            height={dimensions.height - HEADER_SIZE}
            rowCount={GRID_SIZE}
            rowHeight={HEADER_SIZE}
            width={HEADER_SIZE}
            style={{ overflow: "hidden" }}
          >
            {RowHeader}
          </Grid>
        </div>
        <div className="main-grid" ref={mainGridOuterRef}>
          <Grid
            ref={mainGridRef}
            className="grid main-grid-inner"
            columnCount={GRID_SIZE}
            columnWidth={CELL_SIZE}
            height={dimensions.height - HEADER_SIZE}
            rowCount={GRID_SIZE}
            rowHeight={HEADER_SIZE}
            width={dimensions.width - HEADER_SIZE}
            onScroll={handleScroll}
          >
            {Cell}
          </Grid>
        </div>
      </div>
    </div>
  );
};

export default Excel;
