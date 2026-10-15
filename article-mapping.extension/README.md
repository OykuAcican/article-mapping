# OVZ Param Mapper – pyRevit extension (OVERZICHT)

## Install
1. Put `ParamFill.extension` in e.g. `C:\PyRevitExtensions\` (not double-nested).
2. pyRevit > Settings > Custom Extension Directories > add `C:\PyRevitExtensions`.
3. Save Settings and Reload. Tab **OVZ Param Mapper** shows the **OVZ Param Mapper** button.

## OVZ Param Mapper
- **Fill parameter** (left): any editable parameter (text, number, integer,
  yes/no). OVZ_Keynote is the default; '(type)' = type parameter (one value per
  type). Numbers are read in project units; yes/no accepts ja/nee, yes/no, 1/0.
- Right: the whole article list (all columns). **Drag value from column** picks
  the column whose value is dragged; it is marked with a triangle and tinted.
  The mapping stores Parameter + Value per rule, so each parameter has its own
  rules. Old mapping files are read as OVZ_Keynote rules.

- Top bar: choose the VMSW article list (.xlsx, sheet "alle delen") and the shared
  mapping file (.csv, put it on the server so the whole team uses the same one).
- Left: model elements grouped by Category / Family / Type / System (abbreviation,
  else system type) / Size. Filter by discipline/category, search, "only empty
  OVZ_Keynote", "only without article". Views, sheets, materials, project info,
  detail items and other non-physical elements are not listed.
- Group by (up to 4 levels, like a schedule): Category, Family, Type, System
  classification, System type, System abbreviation, Shape, Size, Diameter, Width,
  Height. Right-click a group header to select its rows; drop an article on a
  header to give it to the whole group. The mapping is still saved per
  Category/Family/Type/System/Size, whatever grouping you use.
- Right: the article list with search and discipline filter.
- Select rows on the left, then drag an article onto them (or double-click the
  article, or use "Give to selected rows").
- Apply to model writes OVZ_Keynote on every element of those rows (one Ctrl+Z)
  and saves your choices into the shared mapping.
- Show check colours: in your own 3D view "OVZ check - <user>" (created from the
  active 3D view, without view template) the rows shown in the list are coloured:
  red = parameter empty, orange = value differs from the mapped/new value,
  green (optional) = correct. Clear colours removes them. Run again after changes.
- Find unfilled shows every row with elements whose OVZ_Keynote is empty.
- Next project: open the same mapping file; rows with exactly the same
  Category/Family/Type/System/Size are pre-filled (blue) – check and Apply.

Shared mapping: on save the file is read again and only your changes are merged.
If a colleague changed the same rule in the meantime, you choose whose version
stays. A backup of the previous version is kept as `.csv.bak`.
