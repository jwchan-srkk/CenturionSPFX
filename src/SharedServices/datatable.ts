import * as $ from 'jquery';
import 'datatables.net'; // Import DataTables jQuery plugin


export interface IDataTableOptions {
  selector: string;
  data: any[];
  columns: Array<{ title: string; data: string; width?: string; orderable?: boolean; searchable?: boolean; className?: string; render?: (data: any, type: any, row: any) => any }>;
  order?: Array<[number, 'asc' | 'desc']>;
  pageLength?: number;
  searching?: boolean;
  ordering?: boolean;
  info?: boolean;
  lengthChange?: boolean;
  language?: any;
  layout?: any;
}

export function initDataTable(options: IDataTableOptions): void {
  const $table = $(options.selector);

  // Destroy existing DataTable if exists
  if ($.fn.dataTable.isDataTable($table)) {
    $table.DataTable().destroy();
  }

  const dtConfig: any = {
    data: options.data,
    columns: options.columns,
    order: options.order ?? [],
    pageLength: options.pageLength ?? 10,
    searching: options.searching ?? false,
    ordering: options.ordering ?? true,
    autoWidth: false,
    orderCellsTop: true,
    info: true,
    lengthChange: options.lengthChange ?? true,
    language: options.language ?? {
      emptyTable: "No data available."
    }
  };

  // Only add layout if explicitly provided
  if (options.layout !== undefined) {
    dtConfig.layout = options.layout;
  }

  $table.DataTable(dtConfig);
}

export function reloadDataTable(selector: string, newData: any[]): void {
  const table = $(selector).DataTable();
  table.clear();
  table.rows.add(newData);
  table.draw();
}

// Example utility: Attach a click handler to a button in the table
export function attachActionHandler(
  selector: string,
  actionClass: string,
  handler: (rowData: any) => void
): void {
  $(selector).on("click", actionClass, function () {
    const table = $(selector).DataTable();
    const rowData = table.row($(this).closest("tr")).data();
    handler(rowData);
  });
}

export function initColumnFilters(selector: string): void {
  const table = $(selector).DataTable();

  $(selector + ' thead tr.filters').remove();

  $(selector + ' thead tr').clone(false).addClass('filters').appendTo(selector + ' thead');

  $(selector + ' thead tr:eq(1) th').each(function (i) {
    $(this).html('<input type="text" placeholder="Filter" style="width: 100%; box-sizing: border-box; padding: 5px;" />');

    $('input', this).on('keyup change', function () {
      if (table.column(i).search() !== (this as HTMLInputElement).value) {
        table.column(i).search((this as HTMLInputElement).value).draw();
      }
    }).on('click', function (e) {
      e.stopPropagation();
    }).on('keydown', function (e) {
      // Allow Ctrl+A (or Cmd+A on Mac) to select all text
      if ((e.ctrlKey || e.metaKey) && e.key === 'a') {
        e.stopPropagation();
        // Let the default behavior happen (select all text)
        return true;
      }
      e.stopPropagation();
    });
  });
}

export function clearColumnFilters(selector: string): void {
  const table = $(selector).DataTable();

  $(selector + ' thead tr.filters input').val('');

  (table.columns().search('') as any).draw();
}
