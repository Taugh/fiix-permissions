# Fiix permissions project

Everything needed to regenerate IREN's Fiix v5 permission workbooks. Created 2026-10-02.

## What is in this folder

| File or folder | What it is |
| --- | --- |
| `Fiix_v5_Group_Build_Sheet.xlsx` | **Main deliverable.** The eight management groups as Fiix v5 settings, build order, gaps and decisions. |
| `Fiix_Permissions_Reference.xlsx` | Fiix's documented default permissions, dependencies and descriptions. Background for the build sheet. |
| `build_group_design.py` | Builds the group build sheet. **The group design lives here.** |
| `build_workbook.py` | Builds the reference workbook. |
| `rebuild.ps1` | One command that runs both builders, recalculates in Excel, and saves here. |
| `data/` | Fiix tables copied from the Help Center on 2026-10-02 and checked against the live pages. |
| `source/` | Copy of management's `IREN_Fiix_Enterprise_Permission_Model_Updated.xlsx`. |
| `../fiix-api/` | Separate tool for querying the Fiix API (`fiix_query.py`). |

## Rebuild

```powershell
.\rebuild.ps1                  # both workbooks
.\rebuild.ps1 -What groups     # only the group build sheet
```

Needs Python 3 with `openpyxl` (`python -m pip install openpyxl`) and desktop Excel. Close the workbooks in Excel first.
`rebuild.ps1` fails loudly if any formula shows an error. OneDrive also keeps version history of the two workbooks.

## How to change the design

The group design is a set of tables in `build_group_design.py`. Each permission has an **entry group**, the lowest group that needs it. Groups above it inherit it.

| Letter | Group | Letter | Group |
| --- | --- | --- | --- |
| `G` | General / View-Only | `I` | Inventory Team |
| `T` | Technician | `A` | CMMS Administrator |
| `P` | Maintenance Planner | `S` | Asset Management |
| `M` | Maintenance Planning Manager | `O` | Owner |

`M+I` means either group is an entry point. The chain is `G > T > P > M` and `T > I`, with `A` taking both `M` and `I`.

| To change | Edit |
| --- | --- |
| Who can do something on a page | `PAGES` (view, create/edit, delete, export, import), or `OVERRIDE` for one permission |
| A page not in Fiix's default tables | `EXTRAS` |
| A left-menu item | `MENU` |
| Purchase order steps | `PURCH_TIER` |
| API Add/Find/Change/Remove for a class | `API` |
| Explanation text on a row | `NOTES` |
| Gaps and decisions list | `GAPS` |
| How a matrix row maps to Fiix | `MAP` (keyed by the matrix row number) |

After editing, run `.\rebuild.ps1 -What groups`. The script stops with a clear message if a page, class or matrix row has no entry.

### If management sends a new matrix

1. Save it in `source/` and update `$Mgmt` in `rebuild.ps1`.
2. Run the rebuild. It lists any matrix rows that have no entry in `MAP`.
3. Add those rows to `MAP` and adjust the tiers they imply.

## Decisions built in, and the line to change if they flip

| Decision | Built as | To flip |
| --- | --- | --- |
| Bulk inventory import | Inventory Team (matrix row 118) | Change the import tier from `"I"` to `"A"` on inventory pages, for example `Current Stock` and `Stock` in `PAGES` |
| Planning Manager can delete work orders | Yes (matrix row 23) | `PAGES["Work Orders"]` delete tier `"M"` to `"A"`, and `API["WorkOrder"]` remove tier |
| CMMS Administrator has Inventory rights | Yes (matrix rows 62 to 78) | `PARENTS["A"]` from `["M", "I"]` to `["M"]` |
| Who approves purchase orders | Asset Management and Owner | `PURCH_TIER` entries for Approve and Reject |
| Account Settings (licences) | Owner only | `MENU["Can See 'Account Settings' Menu Item"]` |
| Mobile request submission for General | Off | `API["WorkOrder"]` add tier from `"P"` to `"G"` (also lets General create formal work orders) |

## Sources and scope

- Fiix v5 only. v6 system roles and v6 work orders do not apply. `data/sysroles.txt` and `data/wo_v6.txt` are kept for reference on the grey tab of the reference workbook.
- Fiix Help Center, Permission library: https://helpdesk.fiixsoftware.com/hc/en-us/sections/360010366012-Permission-library
- Default tables were verified against the live pages by hash on 2026-10-02. Fiix changes its docs occasionally; recheck before relying on this after a few months.

## Known limits

- Fiix permissions allow or deny actions per page. They cannot enforce approvals, draft versus active PMs, or field-level limits. See the Requirement Map and Gaps sheets.
- Rows marked Low confidence in `Page Settings` are inferred. Test with one user per group in the Enterprise sandbox before production.
- Python run from the Claude desktop session cannot create files inside this OneDrive folder. `rebuild.ps1` works around it by building in a temporary folder and saving through Excel. In your own terminal, plain `python build_group_design.py <out.xlsx> <source workbook>` also works.
