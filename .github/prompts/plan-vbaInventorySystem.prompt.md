## Plan: Excel Inventory VBA System

TL;DR: This should be built as a single Excel workbook with a strict master-data layer, transaction entry sheets, and centralized VBA routines that auto-save rows, recalculate inventory, and freeze profit snapshots at the moment of sale. The design below follows your requirement that product/supplier/customer names be selected from master data, units auto-fill from item data, and stock/profit summaries update automatically without allowing manual overrides.

### 1. Workbook structure
Create these sheets:

- 基础信息
  - 商品信息
  - 供应商信息
  - 客户信息
- 入库单
- 销售单
- 销售退货单
- 进货退货单
- 报损单
- 报溢单
- 库存统计
- 利润统计
- Optional: 设置 / 导航 / 日志

The master-data sheets should use Excel Tables so VBA can reference them consistently, and the transaction sheets should each have a clear “entry area” plus a separate detail table for saved rows.

### 2. Master-data rules and validation
This is the key business rule layer.

- 商品名称、供应商名称、客户名称 must be selected from dropdowns only.
- Duplicate names within 基础信息 should trigger a message and block save.
- Unit must be filled automatically after product selection.
- If a product name is updated in the base data, all related rows in transaction tables should follow the updated name if your design chooses to maintain the logical item mapping.
- Use Data Validation + VBA validation on save to enforce:
  - no manual typing in restricted fields
  - no duplicate base records
  - no empty required fields
  - consistent numeric validation for quantity/price

### 3. Transaction entry design
For each transaction type, use a standard entry pattern:

- Header fields:
  - 单据号 / 日期 / 供货单位 / 客户 / 商品名称 / 数量 / 单位 / 单价 / 金额（if applicable）
- Save button logic:
  1. validate fields
  2. check the selected product exists in master data
  3. compute amount when required
  4. append the record to the corresponding detail table
  5. clear the form
  6. trigger inventory/profit recalculation

This should be implemented in a reusable VBA module and reused by all transaction sheets.

### 4. Inventory calculation logic
The stock summary should be driven by VBA after each successful save.

Rule:
- 库存数量 = 入库数量 - 销售数量 + 销售退货数量 - 进货退货数量 + 报溢数量 - 报损数量

Average cost rule:
- For incoming stock:
  - 成本均价 = (现有库存总价 + 本次进货总价) ÷ (现有库存数量 + 进货数量)
- For purchase return:
  - 成本均价 = (现有库存总价 - 退货总金额) ÷ (现有库存数量 - 退货数量)

Implementation approach:
- Maintain a per-product stock ledger in 库存统计
- On each document save, update:
  - 商品名称
  - 当前数量
  - 单位
  - 成本均价
  - 库存总价
- Use a central routine like RecalcInventory() and call it after save for all transaction types

### 5. Profit calculation logic
Profit should only use sales and sales return details.

Rule:
- 利润 = 数量 × (销售单价 - 成本均价) + 调整项
- 调整项 defaults to 0
- Snapshot the sales price and cost price at the time of save
- After save, do not allow historical profit rows to change

Implementation approach:
- On each sale or sales return save:
  - copy current product price and current average cost into the profit summary row
  - store the quantity
  - store adjustment
  - compute profit and save it
- The “profit summary” should not be recalculated from current stock data later; it should remain a fixed historical record

### 6. Cross-sheet synchronization
If the basic data is edited later, you should decide whether to:
- update all related display labels / dropdowns in transaction sheets
- keep historical detail rows unchanged
- update only the master reference names and relevant display cells

This matters because your requirement says:
- “如果基础信息中的信息修改，对应所有表格中的信息跟着修改。”

To satisfy that:
- Use a central name-lookup routine
- Trigger a workbook or sheet event when master data changes
- Refresh transaction dropdowns and display values in all relevant areas

### 7. UI and user experience
- Add buttons such as:
  - 保存
  - 清空
  - 关闭
  - 刷新下拉
- Use clear labels and a clean layout
- Put summary sheets in presentation-friendly format with totals and filters

### 8. Suggested VBA module breakdown
Create separate modules with clear responsibilities:

- modBaseData
  - master-data validation
  - duplicate checks
  - list population
- modValidation
  - required field checks
  - numeric validation
  - dropdown enforcement
- modTransactionEntry
  - save logic for each document type
  - clear form behavior
- modInventoryCalc
  - stock quantity formula
  - average cost formula
  - summary update routine
- modProfitCalc
  - sales-profit calculation
  - sales-return handling
  - snapshot logic
- modWorkbookEvents
  - worksheet change events
  - master-data synchronization
- modUI
  - button wiring
  - refresh actions

---

## Relevant files
These are the workbooks/sheets and VBA areas that are relevant to implementation:

- Workbook sheets:
  - 基础信息
  - 入库单
  - 销售单
  - 销售退货单
  - 进货退货单
  - 报损单
  - 报溢单
  - 库存统计
  - 利润统计
- VBA modules to create:
  - modBaseData
  - modValidation
  - modTransactionEntry
  - modInventoryCalc
  - modProfitCalc
  - modWorkbookEvents
  - modUI

---

## Verification
The implementation should be validated with these explicit tests:

1. Duplicate prevention
   - Try adding a duplicate supplier/customer/product name in 基础信息 and confirm save is blocked.

2. Entry-save flow
   - Enter a purchase record, save it, and confirm:
     - the detail row is appended
     - the data entry area clears
     - the item remains selectable in the next row

3. Inventory accuracy
   - Create a receipt, sale, return, loss, and overflow scenario and verify:
     - quantity updates match the formula
     - average cost updates match the stated formula
     - stock summary stays consistent

4. Profit snapshot behavior
   - Save a sale with a price and cost, then change the product master data and verify the historical profit row remains unchanged.

5. Data integrity
   - Confirm that all unit and customer/product names used in the sales and inventory summary still match the base data.

---

## Decision summary
- Use a single workbook with Excel Tables and VBA macros.
- Keep each transaction type in a separate sheet and detail table.
- Centralize inventory and profit calculations in VBA.
- Treat profit as a historical snapshot, not a live recalculation.

This is the cleanest way to satisfy your business rules and avoid incorrect inventory math.

If you want, I can continue and give you the concrete VBA code skeleton next: the workbook structure, tables, and the first macro set for master data validation and saving a receipt entry.
