local function _normalize(s)
  return (s or ""):gsub("%s+", "")
end

local function _has_summary_column(tbl)
  if not tbl.head or not tbl.head.rows or #tbl.head.rows == 0 then
    return false
  end

  local first_row = tbl.head.rows[1]
  if not first_row or not first_row.cells then
    return false
  end

  for i = 1, #first_row.cells do
    local cell = first_row.cells[i]
    local text = ""
    if cell and cell.contents then
      text = pandoc.utils.stringify(cell.contents)
    else
      text = pandoc.utils.stringify(cell)
    end
    if _normalize(text) == "超期内容概括" then
      return true
    end
  end

  return false
end

function Table(tbl)
  if not _has_summary_column(tbl) then
    return tbl
  end

  local n = #tbl.colspecs
  if n ~= 3 then
    return tbl
  end

  -- Keep summary tables clearly narrower than full page width for better PDF layout.
  tbl.colspecs[1][2] = 0.10
  tbl.colspecs[2][2] = 0.07
  tbl.colspecs[3][2] = 0.55
  return tbl
end
