/**
 * Copyright 2023-present DreamNum Co., Ltd.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

import type { IDisposable, IRange, ISelectionCell, Nullable } from '@univerjs/core';
import type {
    IColumnsHeaderCfgParam,
    IRowsHeaderCfgParam,
    RenderComponentType,
    SpreadsheetColumnHeader,
    SpreadsheetRowHeader,
    SpreadsheetSkeleton,
} from '@univerjs/engine-render';

import type { ISelectionStyle } from '@univerjs/sheets';
import type { IScrollState } from '@univerjs/sheets-ui';
import type { FRange } from '@univerjs/sheets/facade';
import { ICommandService, toDisposable } from '@univerjs/core';
import { IRenderManagerService, SHEET_VIEWPORT_KEY, Vector2 } from '@univerjs/engine-render';

import { SetWorksheetRowIsAutoHeightCommand } from '@univerjs/sheets';
import {
    IMarkSelectionService,
    SetColumnHeaderHeightCommand,
    SetRowHeaderWidthCommand,
    SetWorksheetColAutoWidthCommand,
    SetZoomRatioCommand,
    SHEET_VIEW_KEY,
    SheetScrollManagerService,
    SheetSkeletonManagerService,
    SheetsScrollRenderController,
} from '@univerjs/sheets-ui';
import { FWorksheet } from '@univerjs/sheets/facade';

/**
 * @ignore
 */
export interface IFWorksheetUIMixin {
    /**
     * Refresh the canvas.
     * @returns {FWorksheet} The FWorksheet instance for chaining.
     * @example
     * ```ts
     * const fWorkbook = univerAPI.getActiveWorkbook();
     * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
     * if (!fWorksheet) return;
     * fWorksheet.refreshCanvas();
     * ```
     */
    refreshCanvas(): FWorksheet;

    /**
     * Highlight multiple ranges on the worksheet.
     * @param {FRange[]} ranges  The ranges to highlight.
     * @param {Nullable<Partial<ISelectionStyle>>} style - style for highlight ranges.
     * @param {Nullable<ISelectionCell>} primary - primary cell for highlight ranges.
     * @return {IDisposable} An IDisposable to remove the highlights.
     * @example
     * ```ts
     * const fWorkbook = univerAPI.getActiveWorkbook();
     * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
     * if (!fWorksheet) return;
     * const ranges = [fWorksheet.getRange('A1:B2'), fWorksheet.getRange('D4:E5')];
     * const disposable = fWorksheet.highlightRanges(ranges, { fill: 'yellow' });
     *
     * // To remove the highlights later
     * disposable.dispose();
     * ```
     */
    highlightRanges(ranges: FRange[], style?: Nullable<Partial<ISelectionStyle>>, primary?: Nullable<ISelectionCell>): IDisposable;

    /**
     * Set zoom ratio of the worksheet.
     * @param {number} zoomRatio The zoom ratio to set.It should be in the range of 0.1 to 4.0.
     * @returns {FWorksheet} The FWorksheet instance for chaining.
     * @example
     * ```ts
     * const fWorkbook = univerAPI.getActiveWorkbook();
     * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
     * if (!fWorksheet) return;
     *
     * // Set zoom ratio to 200%
     * fWorksheet.zoom(2);
     * const zoomRatio = fWorksheet.getZoom();
     * console.log(zoomRatio); // 2
     * ```
     */
    zoom(zoomRatio: number): FWorksheet;

    /**
     * Get the zoom ratio of the worksheet.
     * @returns {number} The zoom ratio of the worksheet.
     * @example
     * ```ts
     * const fWorkbook = univerAPI.getActiveWorkbook();
     * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
     * if (!fWorksheet) return;
     * const zoomRatio = fWorksheet.getZoom();
     * console.log(zoomRatio);
     * ```
     */
    getZoom(): number;

    /**
     * Get visible range of main viewport.
     * @returns {IRange} - visible range
     * @example
     * ```ts
     * const fWorkbook = univerAPI.getActiveWorkbook();
     * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
     * if (!fWorksheet) return;
     * const visibleRange = fWorksheet.getVisibleRange();
     * console.log(visibleRange);
     * console.log(fWorksheet.getRange(visibleRange).getA1Notation());
     * ```
     */
    getVisibleRange(): IRange | null;

    /**
     * Get visible ranges of all viewports.
     * @returns {Record<SHEET_VIEWPORT_KEY, IRange>} - visible ranges of all viewports
     * @example
     * ```ts
     * const fWorkbook = univerAPI.getActiveWorkbook();
     * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
     * if (!fWorksheet) return;
     * const visibleRanges = fWorksheet.getVisibleRangesOfAllViewports();
     * console.log(visibleRanges);
     * const mainLeftTopViewportRange = visibleRanges?.get(univerAPI.Enum.SHEET_VIEWPORT_KEY.VIEW_MAIN_LEFT_TOP);
     * console.log(fWorksheet.getRange(mainLeftTopViewportRange).getA1Notation());
     * ```
     */
    getVisibleRangesOfAllViewports(): Map<SHEET_VIEWPORT_KEY, IRange> | null;

    /**
     * Scroll spreadsheet(viewMain) to cell position. Make the cell at topleft of current viewport.
     * Based on the limitations of viewport and the number of rows and columns, you can only scroll to the maximum scrollable range.
     * @param {number} row - Cell row index
     * @param {number} column - Cell column index
     * @param {number} [duration] - The duration of the scroll animation in milliseconds.
     * @returns {FWorksheet} - The FWorksheet instance for chaining.
     * @example
     * ```ts
     * const fWorkbook = univerAPI.getActiveWorkbook();
     * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
     * if (!fWorksheet) return;
     *
     * // Scroll to cell D10
     * const fRange = fWorksheet.getRange('D10');
     * const row = fRange.getRow();
     * const column = fRange.getColumn();
     * fWorksheet.scrollToCell(row, column);
     * ```
     */
    scrollToCell(row: number, column: number, duration?: number): FWorksheet;

    /**
     * Get scroll state of current sheet.
     * @returns {IScrollState} curr scroll state
     * @example
     * ```ts
     * const fWorkbook = univerAPI.getActiveWorkbook();
     * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
     * if (!fWorksheet) return;
     *
     * // Scroll to cell D10
     * const fRange = fWorksheet.getRange('D10');
     * const row = fRange.getRow();
     * const column = fRange.getColumn();
     * fWorksheet.scrollToCell(row, column);
     *
     * // Get scroll state
     * const scrollState = fWorksheet.getScrollState();
     * const { offsetX, offsetY, sheetViewStartColumn, sheetViewStartRow } = scrollState;
     * console.log(scrollState); // sheetViewStartRow: 9, sheetViewStartColumn: 3, offsetX: 0, offsetY: 0
     * ```
     */
    getScrollState(): IScrollState;

    /**
     * Get the skeleton service of the worksheet.
     * @returns {Nullable<SpreadsheetSkeleton>} The skeleton of the worksheet.
     * @example
     * ```ts
     * const fWorkbook = univerAPI.getActiveWorkbook();
     * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
     * if (!fWorksheet) return;
     * const skeleton = fWorksheet.getSkeleton();
     * console.log(skeleton);
     * ```
     */
    getSkeleton(): Nullable<SpreadsheetSkeleton>;

    /**
     * Sets the width of the given column to fit its contents.
     * @param {number} columnPosition - The position of the given column to resize. index starts at 0.
     * @returns {FWorksheet} - The FWorksheet instance for chaining.
     * @example
     * ```ts
     * const fWorkbook = univerAPI.getActiveWorkbook();
     * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
     * if (!fWorksheet) return;
     *
     * // Set the long text value in cell A1
     * const fRange = fWorksheet.getRange('A1');
     * fRange.setValue('Whenever it is a damp, drizzly November in my soul...');
     *
     * // Set the column A to a width which fits the text
     * fWorksheet.autoResizeColumn(0);
     * ```
     */
    autoResizeColumn(columnPosition: number): FWorksheet;

    /**
     * Sets the width of all columns starting at the given column position to fit their contents.
     * @param {number} startColumn - The position of the first column to resize. index starts at 0.
     * @param {number} numColumns - The number of columns to auto-resize.
     * @returns {FWorksheet} - The FWorksheet instance for chaining.
     * @example
     * ```ts
     * const fWorkbook = univerAPI.getActiveWorkbook();
     * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
     * if (!fWorksheet) return;
     *
     * // Set the A:C columns to a width that fits their text.
     * fWorksheet.autoResizeColumns(0, 3);
     * ```
     */
    autoResizeColumns(startColumn: number, numColumns: number): FWorksheet;

    /**
     * Sets the width of all columns starting at the given column position to fit their contents.
     * @deprecated use `autoResizeColumns` instead
     * @param {number} columnPosition - The position of the first column to resize. index starts at 0.
     * @param {number} numColumn - The number of columns to auto-resize.
     * @returns {FWorksheet} - The FWorksheet instance for chaining.
     */
    setColumnAutoWidth(columnPosition: number, numColumn: number): FWorksheet;

    /**
     * Sets the height of all rows starting at the given row position to fit their contents.
     * @param {number} startRow - The position of the first row to resize. index starts at 0.
     * @param {number} numRows - The number of rows to auto-resize.
     * @returns {FWorksheet} - The FWorksheet instance for chaining.
     * @example
     * ```ts
     * const fWorkbook = univerAPI.getActiveWorkbook();
     * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
     * if (!fWorksheet) return;
     *
     * // Set the first 3 rows to a height that fits their text.
     * fWorksheet.autoResizeRows(0, 3);
     * ```
     */
    autoResizeRows(startRow: number, numRows: number): FWorksheet;

    /**
     * Customize the column header of the worksheet.
     * @param {IColumnsHeaderCfgParam} cfg The configuration of the column header.
     * @example
     * ```typescript
     * const fWorkbook = univerAPI.getActiveWorkbook();
     * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
     * if (!fWorksheet) return;
     * fWorksheet.customizeColumnHeader({
     *   headerStyle: {
     *     fontColor: '#fff',
     *     backgroundColor: '#4e69ee',
     *     fontSize: 9
     *   },
     *   columnsCfg: {
     *     0: 'kuma II',
     *     3: {
     *       text: 'Size',
     *       textAlign: 'left', // CanvasTextAlign
     *       fontColor: '#fff',
     *       fontSize: 12,
     *       borderColor: 'pink',
     *       backgroundColor: 'pink',
     *     },
     *     4: 'Wow'
     *   }
     * });
     * ```
     */
    customizeColumnHeader(cfg: IColumnsHeaderCfgParam): void;

    /**
     * Customize the row header of the worksheet.
     * @param {IRowsHeaderCfgParam} cfg The configuration of the row header.
     * @example
     * ```typescript
     * const fWorkbook = univerAPI.getActiveWorkbook();
     * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
     * if (!fWorksheet) return;
     * fWorksheet.customizeRowHeader({
     *   headerStyle: {
     *     backgroundColor: 'pink',
     *     fontSize: 12
     *   },
     *   rowsCfg: {
     *     0: 'Moka II',
     *     3: {
     *       text: 'Size',
     *       textAlign: 'left', // CanvasTextAlign
     *     },
     *   }
     * });
     * ```
     */
    customizeRowHeader(cfg: IRowsHeaderCfgParam): void;

    /**
     * Set column height for column header.
     * @param {number} height - The height to set.
     * @returns {FWorksheet} - The FWorksheet instance for chaining.
     * @example
     * ```ts
     * const fWorkbook = univerAPI.getActiveWorkbook();
     * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
     * if (!fWorksheet) return;
     * fWorksheet.setColumnHeaderHeight(100);
     * ```
     */
    setColumnHeaderHeight(height: number): FWorksheet;

    /**
     * Set column height for column header.
     * @param {number} width - The width to set.
     * @returns {FWorksheet} - The FWorksheet instance for chaining.
     * @example
     * ```ts
     * const fWorkbook = univerAPI.getActiveWorkbook();
     * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
     * if (!fWorksheet) return;
     * fWorksheet.setRowHeaderWidth(100);
     * ```
     */
    setRowHeaderWidth(width: number): FWorksheet;

    /**
     * Hit test: get the cell position from client coordinates.
     * @param {number} clientX The x coordinate relative to the viewport.
     * @param {number} clientY The y coordinate relative to the viewport.
     * @returns {Nullable<{ row: number; column: number }>} The cell position or null.
     * @example
     * ```ts
     * const fWorksheet = univerAPI.getActiveWorkbook().getSheetByName('Sheet1');
     * if (!fWorksheet) return;
     * const result = fWorksheet.hitTest(100, 50);
     * if (result) {
     *   console.log(`Cell: row=${result.row}, column=${result.column}`);
     * }
     * ```
     */
    hitTest(clientX: number, clientY: number): Nullable<{ row: number; column: number }>;

    /**
     * Get the pixel rectangle of the specified cell, relative to the canvas element.
     * This is useful for positioning overlay elements (e.g., highlight borders) on the canvas.
     * @param {number} row The row index (0-based).
     * @param {number} column The column index (0-based).
     * @returns {Nullable<{ left: number; top: number; width: number; height: number }>} The pixel rectangle, or null if skeleton is not available.
     * @example
     * ```ts
     * const fWorksheet = univerAPI.getActiveWorkbook().getSheetByName('Sheet1');
     * if (!fWorksheet) return;
     * const rect = fWorksheet.getCellRect(0, 0);
     * if (rect) {
     *   console.log(`Cell(0,0) rect: left=${rect.left}, top=${rect.top}, width=${rect.width}, height=${rect.height}`);
     * }
     * ```
     */
    getCellRect(row: number, column: number): Nullable<{ left: number; top: number; width: number; height: number }>;
}

export class FWorksheetUIMixin extends FWorksheet implements IFWorksheetUIMixin {
    override refreshCanvas(): FWorksheet {
        const renderManagerService = this._injector.get(IRenderManagerService);
        const unitId = this._fWorkbook.id;
        const render = renderManagerService.getRenderUnitById(unitId);

        if (!render) {
            throw new Error(`Render Unit with unitId ${unitId} not found`);
        }

        render.with(SheetSkeletonManagerService).reCalculate();

        render.components.forEach((component) => {
            component.makeDirty?.();
        });

        return this;
    }

    override highlightRanges(ranges: FRange[], style?: Nullable<Partial<ISelectionStyle>>, primary?: Nullable<ISelectionCell>): IDisposable {
        const markSelectionService = this._injector.get(IMarkSelectionService);
        const ids: string[] = [];
        for (const range of ranges) {
            const iRange = range.getRange();
            const id = markSelectionService.addShapeWithNoFresh({ range: iRange, style, primary });
            if (id) {
                ids.push(id);
            }
        }
        markSelectionService.refreshShapes();

        if (ids.length === 0) {
            throw new Error('Failed to highlight current range');
        }
        return toDisposable(() => {
            ids.forEach((id) => {
                markSelectionService.removeShape(id);
            });
        });
    }

    override zoom(zoomRatio: number): FWorksheet {
        const commandService = this._injector.get(ICommandService);
        const _zoomRatio = Math.min(Math.max(zoomRatio, 0.1), 4);
        commandService.executeCommand(SetZoomRatioCommand.id, {
            unitId: this._workbook.getUnitId(),
            subUnitId: this._worksheet.getSheetId(),
            zoomRatio: _zoomRatio,
        });
        return this;
    }

    override getZoom(): number {
        return this._worksheet.getZoomRatio();
    }

    override getVisibleRange(): IRange | null {
        const unitId = this._workbook.getUnitId();
        const renderManagerService = this._injector.get(IRenderManagerService);
        const render = renderManagerService.getRenderUnitById(unitId);
        if (!render) return null;
        const skm = render.with(SheetSkeletonManagerService);
        const sk = skm.getCurrentSkeleton();
        if (!sk) return null;
        return sk.getVisibleRangeByViewport(SHEET_VIEWPORT_KEY.VIEW_MAIN) as IRange;
    }

    override getVisibleRangesOfAllViewports(): Map<SHEET_VIEWPORT_KEY, IRange> | null {
        const unitId = this._workbook.getUnitId();
        const renderManagerService = this._injector.get(IRenderManagerService);
        const render = renderManagerService.getRenderUnitById(unitId);
        if (!render) return null;
        const skm = render.with(SheetSkeletonManagerService);
        const sk = skm.getCurrentSkeleton();
        if (!sk) return null;
        return sk.getVisibleRanges();
    }

    override scrollToCell(row: number, column: number, duration?: number): FWorksheet {
        const unitId = this._workbook.getUnitId();
        const renderManagerService = this._injector.get(IRenderManagerService);
        const render = renderManagerService.getRenderUnitById(unitId);
        if (render) {
            const scrollRenderController = render?.with(SheetsScrollRenderController);
            scrollRenderController.scrollToCell(row, column, duration);
        }
        return this;
    }

    override getScrollState(): IScrollState {
        const emptyScrollState: IScrollState = {
            offsetX: 0,
            offsetY: 0,
            sheetViewStartColumn: 0,
            sheetViewStartRow: 0,
        };
        const unitId = this._workbook.getUnitId();
        const sheetId = this._worksheet.getSheetId();
        const renderManagerService = this._injector.get(IRenderManagerService);
        const render = renderManagerService.getRenderUnitById(unitId);
        if (!render) return emptyScrollState;
        const sheetScrollManagerService = render.with(SheetScrollManagerService);
        const scrollState = sheetScrollManagerService.getScrollStateByParam({ unitId, sheetId });
        return scrollState || emptyScrollState;
    }

    override getSkeleton(): Nullable<SpreadsheetSkeleton> {
        const service = this._injector.get(IRenderManagerService).getRenderUnitById(this._workbook.getUnitId())?.with(SheetSkeletonManagerService);
        return service?.getSkeleton(this._worksheet.getSheetId());
    }

    /**
     * Hit test to get the cell at the given page coordinates.
     * @param clientX - X coordinate relative to the viewport (pageX)
     * @param clientY - Y coordinate relative to the viewport (pageY)
     * @returns The cell row and column at the given coordinates, or null if not found.
     *          Also returns the cell's pixel position (left, top, width, height) relative to the canvas element.
     */
    override hitTest(clientX: number, clientY: number): Nullable<{
        row: number;
        column: number;
        left?: number;
        top?: number;
        width?: number;
        height?: number;
    }> {
        const unitId = this._workbook.getUnitId();
        const renderManagerService = this._injector.get(IRenderManagerService);
        const render = renderManagerService.getRenderUnitById(unitId);

        if (!render) {
            console.warn('[hitTest] Render not found for unitId:', unitId);
            return null;
        }

        const scene = render.scene;
        if (!scene) {
            console.warn('[hitTest] Scene not found');
            return null;
        }

        // ============================================================
        // STEP 1: Convert client coordinates to canvas-relative coordinates
        // monitor.getClientOffset() returns coordinates relative to viewport
        // We need to convert to coordinates relative to the canvas element
        // ============================================================
        // 注意：scene.getEngine().getCanvas() 返回的是 Univer 的 Canvas 包装对象，不是 DOM 元素
        // 需要使用 scene.getEngine().getCanvas().getCanvasElement() 获取实际的 DOM 元素
        const engine = scene.getEngine();
        const canvasElement = engine?.getCanvasElement();
        let offsetX: number;
        let offsetY: number;

        if (canvasElement) {
            const canvasRect = canvasElement.getBoundingClientRect();
            // Client coordinates are relative to viewport, convert to canvas-relative
            offsetX = clientX - canvasRect.left;
            offsetY = clientY - canvasRect.top;
        } else {
            // Fallback: use client coordinates directly
            offsetX = clientX;
            offsetY = clientY;
            console.warn('[hitTest] Canvas element not found, using client coordinates directly');
        }

        // ============================================================
        // STEP 2: Transform coordinates using getCoordRelativeToViewport
        // This is the SAME process as Univer's internal mouse event handling
        // Reference: selection-shape-extension.ts line 282-304
        //
        // IMPORTANT: getCoordRelativeToViewport adds viewport scroll offsets
        // to the coordinates. Then getCellByOffset internally uses scrollXY
        // to calculate scroll offsets AGAIN, causing double offset.
        //
        // SOLUTION: Do NOT use getCoordRelativeToViewport.
        // Instead, pass canvas-relative coordinates directly to getCellByOffset.
        // getCellByOffset internally handles scroll and scale transformations.
        // ============================================================

        // Get scroll and scale information using canvas-relative coordinates
        const coordVector = Vector2.FromArray([offsetX, offsetY]);
        const scrollXY = scene.getScrollXYInfoByViewport(coordVector);
        const { scaleX, scaleY } = scene.getAncestorScale();

        // ============================================================
        // STEP 3: Get skeleton
        // ============================================================
        const skeleton = this.getSkeleton();
        if (!skeleton) {
            console.warn('[hitTest] Skeleton not found');
            return null;
        }

        // ============================================================
        // STEP 4: Call getCellByOffset with canvas-relative coordinates
        //
        // getCellByOffset internally calls getTransformOffsetX/Y which does:
        //   offsetX = offsetX / scaleX + scrollX - rowHeaderWidth
        //   offsetY = offsetY / scaleY + scrollY - columnHeaderHeight
        //
        // So we should pass canvas-relative coordinates (offsetX, offsetY)
        // and the correct scrollXY. getCellByOffset will handle the rest.
        // ============================================================
        const cellInfo = skeleton.getCellByOffset(offsetX, offsetY, scaleX, scaleY, scrollXY);

        if (!cellInfo) {
            console.warn('[hitTest] No cell found at coordinates', { offsetX, offsetY });
            return null;
        }

        // Calculate cell pixel position relative to canvas
        // Use rowHeightAccumulation and columnWidthAccumulation to get cell position
        const row = cellInfo.actualRow;
        const column = cellInfo.actualColumn;

        // Get row height and column width accumulation arrays
        // NOTE: rowHeightAccumulation and columnWidthAccumulation are properties of SheetSkeleton, NOT methods
        const rowHeightAccumulation = (skeleton as any).rowHeightAccumulation as number[] | undefined;
        const columnWidthAccumulation = (skeleton as any).columnWidthAccumulation as number[] | undefined;

        let left = 0;
        let top = 0;
        let width = 0;
        let height = 0;

        if (rowHeightAccumulation && columnWidthAccumulation) {
            // Calculate left position (startX)
            if (column > 0) {
                left = columnWidthAccumulation[column - 1] || 0;
            }

            // Calculate top position (startY)
            if (row > 0) {
                top = rowHeightAccumulation[row - 1] || 0;
            }

            // Calculate width and height
            width = (columnWidthAccumulation[column] || 0) - left;
            height = (rowHeightAccumulation[row] || 0) - top;
        }

        return {
            row,
            column,
            left,
            top,
            width,
            height,
        };
    }

    /**
     * Get the pixel rectangle of the specified cell, relative to the canvas element.
     * This is useful for positioning overlay elements (e.g., highlight borders) on the canvas.
     * @param row - The row index (0-based).
     * @param column - The column index (0-based).
     * @returns The pixel rectangle, or null if skeleton is not available.
     */
    override getCellRect(row: number, column: number): Nullable<{ left: number; top: number; width: number; height: number }> {
        const skeleton = this.getSkeleton();
        if (!skeleton) {
            console.warn('[getCellRect] Skeleton not found');
            return null;
        }

        // NOTE: rowHeightAccumulation and columnWidthAccumulation are properties of SheetSkeleton, NOT methods
        const rowHeightAccumulation = (skeleton as any).rowHeightAccumulation as number[] | undefined;
        const columnWidthAccumulation = (skeleton as any).columnWidthAccumulation as number[] | undefined;

        if (!rowHeightAccumulation || !columnWidthAccumulation) {
            console.warn('[getCellRect] Accumulation arrays not available');
            return null;
        }

        let left = 0;
        let top = 0;

        // Calculate left position (startX)
        if (column > 0) {
            left = columnWidthAccumulation[column - 1] || 0;
        }

        // Calculate top position (startY)
        if (row > 0) {
            top = rowHeightAccumulation[row - 1] || 0;
        }

        // Calculate width and height
        const width = (columnWidthAccumulation[column] || 0) - left;
        const height = (rowHeightAccumulation[row] || 0) - top;

        // Add row header width and column header height offset
        // The accumulation arrays are relative to the sheet content area (without headers).
        // The canvas rendering includes row/column headers, so we need to add these offsets
        // to get the position relative to the canvas container.
        // Use rowHeaderWidthAndMarginLeft / columnHeaderHeightAndMarginTop to include margin offsets.
        const offsetX = (skeleton as any).rowHeaderWidthAndMarginLeft || 0;
        const offsetY = (skeleton as any).columnHeaderHeightAndMarginTop || 0;

        return {
            left: left + offsetX,
            top: top + offsetY,
            width,
            height,
        };
    }

    override autoResizeColumn(columnPosition: number): FWorksheet {
        return this.autoResizeColumns(columnPosition, 1);
    }

    override autoResizeColumns(startColumn: number, numColumns: number): FWorksheet {
        const unitId = this._workbook.getUnitId();
        const subUnitId = this._worksheet.getSheetId();
        const ranges = [
            {
                startColumn,
                endColumn: startColumn + numColumns - 1,
                startRow: 0,
                endRow: this._worksheet.getRowCount() - 1,
            },
        ];

        this._commandService.syncExecuteCommand(SetWorksheetColAutoWidthCommand.id, {
            unitId,
            subUnitId,
            ranges,
        });

        return this;
    }

    override setColumnAutoWidth(columnPosition: number, numColumn: number): FWorksheet {
        return this.autoResizeColumns(columnPosition, numColumn);
    }

    override autoResizeRows(startRow: number, numRows: number): FWorksheet {
        const unitId = this._workbook.getUnitId();
        const subUnitId = this._worksheet.getSheetId();
        const ranges = [
            {
                startRow,
                endRow: startRow + numRows - 1,
                startColumn: 0,
                endColumn: this._worksheet.getColumnCount() - 1,
            },
        ];

        this._commandService.syncExecuteCommand(SetWorksheetRowIsAutoHeightCommand.id, {
            unitId,
            subUnitId,
            ranges,
        });

        return this;
    }

    override customizeColumnHeader(cfg: IColumnsHeaderCfgParam): void {
        const unitId = this._workbook.getUnitId();
        const subUnitId = this._worksheet.getSheetId();

        const renderManagerService = this._injector.get(IRenderManagerService);
        const render = renderManagerService.getRenderUnitById(unitId);
        if (render && cfg.headerStyle?.size) {
            const skm = render.with(SheetSkeletonManagerService);
            skm.setColumnHeaderSize(render, subUnitId, cfg.headerStyle?.size);
        }

        const sheetColumn = this._getSheetRenderComponent(unitId, SHEET_VIEW_KEY.COLUMN) as SpreadsheetColumnHeader;
        sheetColumn.setCustomHeader(cfg, subUnitId);
    }

    override customizeRowHeader(cfg: IRowsHeaderCfgParam): void {
        const unitId = this._workbook.getUnitId();
        const subUnitId = this._worksheet.getSheetId();

        const renderManagerService = this._injector.get(IRenderManagerService);
        const render = renderManagerService.getRenderUnitById(unitId);
        if (render && cfg.headerStyle?.size) {
            const skm = render.with(SheetSkeletonManagerService);
            skm.setRowHeaderSize(render, subUnitId, cfg.headerStyle?.size);
        }

        const sheetRow = this._getSheetRenderComponent(unitId, SHEET_VIEW_KEY.ROW) as SpreadsheetRowHeader;
        sheetRow.setCustomHeader(cfg, subUnitId);
    }

    override setColumnHeaderHeight(height: number): FWorksheet {
        const unitId = this._workbook.getUnitId();
        const subUnitId = this._worksheet.getSheetId();

        this._commandService.executeCommand(SetColumnHeaderHeightCommand.id, {
            unitId,
            subUnitId,
            size: height,
        });
        return this;
    }

    override setRowHeaderWidth(width: number): FWorksheet {
        const unitId = this._workbook.getUnitId();
        const subUnitId = this._worksheet.getSheetId();

        this._commandService.executeCommand(SetRowHeaderWidthCommand.id, {
            unitId,
            subUnitId,
            size: width,
        });
        return this;
    }

    /**
     * Get sheet render component from render by unitId and view key.
     * @private
     * @param {string} unitId The unit id of the spreadsheet.
     * @param {SHEET_VIEW_KEY} viewKey The view key of the spreadsheet.
     * @returns {Nullable<RenderComponentType>} The render component.
     */
    private _getSheetRenderComponent(unitId: string, viewKey: SHEET_VIEW_KEY): Nullable<RenderComponentType> {
        const renderManagerService = this._injector.get(IRenderManagerService);
        const render = renderManagerService.getRenderUnitById(unitId);
        if (!render) {
            throw new Error(`Render Unit with unitId ${unitId} not found`);
        }

        const { components } = render;
        const renderComponent = components.get(viewKey);
        if (!renderComponent) {
            throw new Error('Render component not found');
        }

        return renderComponent;
    }
}

FWorksheet.extend(FWorksheetUIMixin);
declare module '@univerjs/sheets/facade' {
    // eslint-disable-next-line ts/naming-convention
    interface FWorksheet extends IFWorksheetUIMixin { }
}
