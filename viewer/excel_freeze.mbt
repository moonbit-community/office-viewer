///|
#cfg(target="js")
extern "js" fn install_excel_freeze(surface_id : String) =
  #| id => {
  #|   const host = document.getElementById(id);
  #|   if (!host?.classList.contains('sheet-surface')) return;
  #|   host.__officeFreezeCleanup?.();
  #|   const cleanups = [], sections = [];
  #|   let freezeSheets;
  #|   try { freezeSheets = JSON.parse(host.querySelector('.office-freeze-metadata')?.textContent || '{}').sheets; } catch {}
  #|   const book = host.querySelector('.sheet-document');
  #|   if (book) book.style.minWidth = '0';
  #|   for (const section of host.querySelectorAll('section[data-sheet]')) {
  #|     const table = section.querySelector('table');
  #|     if (!table) continue;
  #|     let viewport = section.querySelector('.office-sheet-viewport');
  #|     if (!viewport) {
  #|       viewport = document.createElement('div');
  #|       viewport.className = 'office-sheet-viewport';
  #|       Object.assign(viewport.style, { overflow: 'auto', maxHeight: '65vh', maxWidth: '100%', position: 'relative' });
  #|       const grid = document.createElement('div');
  #|       grid.className = 'office-sheet-grid';
  #|       Object.assign(grid.style, { position: 'relative', width: 'max-content', minWidth: '100%' });
  #|       table.parentNode.insertBefore(viewport, table);
  #|       viewport.appendChild(grid); grid.appendChild(table);
  #|       const objects = section.querySelector(':scope > .sheet-objects');
  #|       if (objects) grid.appendChild(objects);
  #|     }
  #|     // Object layout runs once before wrapping and sizes the section.
  #|     // The grid now owns those dimensions so the viewport can scroll.
  #|     section.style.removeProperty('min-width');
  #|     section.style.removeProperty('min-height');
  #|     const count = name => {
  #|       const metadata = freezeSheets && Object.hasOwn(freezeSheets, section.dataset.sheet) ? freezeSheets[section.dataset.sheet] : null;
  #|       const value = Number(metadata ? metadata[name === 'freezeRows' ? 'rows' : 'columns'] : section.dataset[name]);
  #|       return Number.isSafeInteger(value) && value >= 0 ? value : 0;
  #|     };
  #|     sections.push({ section, table, rows: count('freezeRows'), columns: count('freezeColumns') });
  #|     const scroll = () => host.__officeLayoutAnnotations?.();
  #|     viewport.addEventListener('scroll', scroll);
  #|     cleanups.push(() => viewport.removeEventListener('scroll', scroll));
  #|   }
  #|   const layout = () => {
  #|     for (const { section, table, rows, columns } of sections) {
  #|       if (section.hidden || !table.offsetWidth) continue;
  #|       const headings = [...table.querySelectorAll('thead th')];
  #|       const bodyRows = [...table.querySelectorAll('tbody tr')];
  #|       const left = [], top = [];
  #|       let x = 0, y = table.tHead?.offsetHeight || 0;
  #|       for (const heading of headings) { left.push(x); x += heading.offsetWidth; }
  #|       for (const row of bodyRows) { top.push(y); y += row.offsetHeight; }
  #|       const pin = (cell, row, column, heading) => {
  #|         const frozenRow = heading || row <= rows;
  #|         const frozenColumn = column === 0 || column <= columns;
  #|         if (!frozenRow && !frozenColumn) return;
  #|         cell.style.position = 'sticky';
  #|         if (frozenRow) cell.style.top = `${heading ? 0 : top[row - 1] || 0}px`;
  #|         if (frozenColumn) cell.style.left = `${left[column] || 0}px`;
  #|         cell.style.zIndex = frozenRow && frozenColumn ? '7' : frozenRow ? '5' : '4';
  #|         if (getComputedStyle(cell).backgroundColor === 'rgba(0, 0, 0, 0)') cell.style.backgroundColor = '#fff';
  #|       };
  #|       headings.forEach((cell, column) => pin(cell, 0, column, true));
  #|       bodyRows.forEach((row, index) => {
  #|         const rowHeading = row.querySelector('.office-row-heading');
  #|         if (rowHeading) pin(rowHeading, index + 1, 0, false);
  #|         for (const cell of row.querySelectorAll('td[data-row][data-col]'))
  #|           pin(cell, Number(cell.dataset.row), Number(cell.dataset.col), false);
  #|       });
  #|     }
  #|   };
  #|   host.__officeFreezeLayout = layout;
  #|   const observer = typeof ResizeObserver === 'function' ? new ResizeObserver(layout) : null;
  #|   observer?.observe(host);
  #|   cleanups.push(() => observer?.disconnect());
  #|   const interactionCleanup = host.__officeViewerCleanup;
  #|   host.__officeFreezeCleanup = () => {
  #|     for (const cleanup of cleanups) cleanup();
  #|     host.__officeFreezeLayout = null; host.__officeFreezeCleanup = null;
  #|   };
  #|   host.__officeViewerCleanup = () => { host.__officeFreezeCleanup?.(); interactionCleanup?.(); };
  #|   layout();
  #|   host.__officeLayoutObjects?.();
  #| }

///|
#cfg(target="js")
extern "js" fn refresh_excel_freeze(surface_id : String) =
  #| id => document.getElementById(id)?.__officeFreezeLayout?.()
