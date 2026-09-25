import { read, utils, writeFileXLSX } from 'xlsx'

import type {
  ExcelImportRow,
  InventoryRow,
  Movement,
  Sale,
} from '../types/warehouse'

export const EXCEL_HEADERS = [
  'ID prodotto',
  'Prodotto',
  'Fornitore',
  'Tipologia',
  'Prezzo normale (€)',
  'Prezzo confidenziale (€)',
  'Prezzo SOMS (€)',
  'Disponibilità',
  'Unità',
  'Scadenza',
]

const REQUIRED_EXCEL_HEADERS = EXCEL_HEADERS.filter(
  (header) => header !== 'ID prodotto'
)

export function makeDateTimeFileName(prefix: string) {
  const now = new Date()

  const date = [
    now.getFullYear(),
    String(now.getMonth() + 1).padStart(2, '0'),
    String(now.getDate()).padStart(2, '0'),
  ].join('-')

  const time = [
    String(now.getHours()).padStart(2, '0'),
    String(now.getMinutes()).padStart(2, '0'),
  ].join('-')

  return `${prefix}_${date}_${time}.xlsx`
}

function saveExcel(
  rows: Record<string, unknown>[],
  sheetName: string,
  filenamePrefix: string,
  widths?: number[]
) {
  const worksheet = utils.json_to_sheet(rows)

  if (widths) {
    worksheet['!cols'] = widths.map((wch) => ({ wch }))
  }

  const workbook = utils.book_new()
  utils.book_append_sheet(workbook, worksheet, sheetName)

  const filename = makeDateTimeFileName(filenamePrefix)

  writeFileXLSX(workbook, filename, {
    compression: true,
  })

  return filename
}

export function exportInventoryExcel(inventory: InventoryRow[]) {
  const rows = inventory.map((product) => ({
    'ID prodotto': product.product_id,
    Prodotto: product.name,
    Fornitore: product.supplier ?? '',
    Tipologia: product.category ?? '',
    'Prezzo normale (€)': product.normal_price ?? '',
    'Prezzo confidenziale (€)': product.confidential_price ?? '',
    'Prezzo SOMS (€)': product.sale_price ?? '',
    Disponibilità: product.availability,
    Unità: product.unit_type,
    Scadenza: product.expiration_date ?? '',
  }))

  return saveExcel(
    rows,
    'Inventario',
    'Inventario',
    [38, 35, 28, 22, 20, 24, 18, 16, 14, 16]
  )
}

export function exportMovementsExcel(movements: Movement[]) {
  const rows = movements.map((movement) => ({
    Data: new Date(movement.created_at).toLocaleString('it-IT'),
    Prodotto: getProductName(movement.products),
    Movimento: movement.movement_type,
    Quantità: movement.quantity,
    'Stock prima': movement.stock_before,
    'Stock dopo': movement.stock_after,
    'Costo unitario (€)': movement.unit_cost ?? '',
    Note: movement.notes ?? '',
  }))

  return saveExcel(
    rows,
    'Movimenti',
    'Movimenti',
    [22, 35, 25, 14, 14, 14, 20, 45]
  )
}

export function exportSalesExcel(sales: Sale[]) {
  const rows: Record<string, unknown>[] = []

  for (const sale of sales) {
    if (!sale.sale_items?.length) {
      rows.push({
        'ID vendita': sale.id,
        Data: new Date(sale.created_at).toLocaleString('it-IT'),
        Stato: sale.status,
        Pagamento: sale.payment_method ?? '',
        Prodotto: '',
        Quantità: '',
        'Prezzo unitario (€)': '',
        'Totale riga (€)': '',
        'Totale vendita (€)': sale.total_amount,
        'Motivo annullamento': sale.cancellation_reason ?? '',
      })

      continue
    }

    for (const item of sale.sale_items) {
      rows.push({
        'ID vendita': sale.id,
        Data: new Date(sale.created_at).toLocaleString('it-IT'),
        Stato: sale.status,
        Pagamento: sale.payment_method ?? '',
        Prodotto: getProductName(item.products),
        Quantità: item.quantity,
        'Prezzo unitario (€)': item.unit_sale_price,
        'Totale riga (€)': item.line_total,
        'Totale vendita (€)': sale.total_amount,
        'Motivo annullamento': sale.cancellation_reason ?? '',
      })
    }
  }

  return saveExcel(
    rows,
    'Storico vendite',
    'Storico_Vendite',
    [38, 22, 16, 16, 35, 12, 20, 18, 20, 40]
  )
}

export async function parseInventoryExcel(
  file: File,
  currentProducts: InventoryRow[]
): Promise<ExcelImportRow[]> {
  const workbook = read(await file.arrayBuffer(), {
    cellDates: true,
  })

  if (!workbook.SheetNames.length) {
    throw new Error('Il file Excel non contiene fogli')
  }

  const worksheet = workbook.Sheets[workbook.SheetNames[0]]

  const matrix = utils.sheet_to_json<unknown[]>(worksheet, {
    header: 1,
    defval: '',
    raw: true,
  })

  if (!matrix.length) {
    throw new Error('Il file Excel è vuoto')
  }

  const headers = (matrix[0] ?? []).map((value) =>
    String(value).trim()
  )

  const missingHeaders = REQUIRED_EXCEL_HEADERS.filter(
    (header) => !headers.includes(header)
  )

  if (missingHeaders.length) {
    throw new Error(
      `Formato Excel non valido. Colonne mancanti: ${missingHeaders.join(', ')}`
    )
  }

  const objects = utils.sheet_to_json<Record<string, unknown>>(
    worksheet,
    {
      defval: '',
      raw: true,
    }
  )

  const parsedRows: ExcelImportRow[] = []

  objects.forEach((row, index) => {
    const excelRowNumber = index + 2

    const productIdValue = String(
      row['ID prodotto'] ?? ''
    ).trim()

    const name = String(row['Prodotto'] ?? '').trim()

    if (!name && !productIdValue) {
      return
    }

    if (!name) {
      throw new Error(
        `Riga ${excelRowNumber}: prodotto mancante`
      )
    }

    const supplier = String(
      row['Fornitore'] ?? ''
    ).trim()

    const category = String(
      row['Tipologia'] ?? ''
    ).trim()

    if (!supplier) {
      throw new Error(
        `Riga ${excelRowNumber}: fornitore mancante`
      )
    }

    if (!category) {
      throw new Error(
        `Riga ${excelRowNumber}: tipologia mancante`
      )
    }

    const availability = parseExcelNumber(
      row['Disponibilità'],
      true
    )

    if (availability === null || availability < 0) {
      throw new Error(
        `Riga ${excelRowNumber}: disponibilità non valida`
      )
    }

    const unitType = normalizeUnit(row['Unità'])
    const expirationDate = parseExcelDate(row['Scadenza'])

    const normalPrice = parseExcelNumber(
      row['Prezzo normale (€)']
    )

    const confidentialPrice = parseExcelNumber(
      row['Prezzo confidenziale (€)']
    )

    const salePrice = parseExcelNumber(
      row['Prezzo SOMS (€)']
    )

    let existing: InventoryRow | undefined

    if (productIdValue) {
      existing = currentProducts.find(
        (product) => product.product_id === productIdValue
      )

      if (!existing) {
        throw new Error(
          `Riga ${excelRowNumber}: ID prodotto non trovato nel database`
        )
      }
    } else {
      existing = currentProducts.find(
        (product) =>
          normalizeText(product.name) === normalizeText(name) &&
          normalizeText(product.supplier ?? '') ===
            normalizeText(supplier)
      )
    }

    const currentAvailability =
      existing?.availability ?? null

    const difference =
      currentAvailability === null
        ? availability
        : availability - currentAvailability

    const changes = existing
      ? buildChanges(existing, {
          name,
          supplier,
          category,
          normal_price: normalPrice,
          confidential_price: confidentialPrice,
          sale_price: salePrice,
          availability,
          unit_type: unitType,
          expiration_date: expirationDate,
        })
      : ['Nuovo prodotto']

    parsedRows.push({
      product_id:
        productIdValue ||
        existing?.product_id ||
        null,

      name,
      supplier,
      category,

      normal_price: normalPrice,
      confidential_price: confidentialPrice,
      sale_price: salePrice,

      availability,
      unit_type: unitType,
      expiration_date: expirationDate,

      currentAvailability,
      difference,

      action: existing ? 'AGGIORNA' : 'NUOVO',
      changes,
    })
  })

  if (!parsedRows.length) {
    throw new Error(
      'Il file non contiene prodotti da importare'
    )
  }

  checkDuplicates(parsedRows)

  return parsedRows
}

function checkDuplicates(rows: ExcelImportRow[]) {
  const keys = rows.map((row) =>
    row.product_id
      ? `id:${row.product_id}`
      : `name:${normalizeText(row.name)}|supplier:${normalizeText(row.supplier)}`
  )

  if (new Set(keys).size !== keys.length) {
    throw new Error(
      'Il file contiene prodotti duplicati'
    )
  }
}

function buildChanges(
  current: InventoryRow,
  next: {
    name: string
    supplier: string
    category: string
    normal_price: number | null
    confidential_price: number | null
    sale_price: number | null
    availability: number
    unit_type: string
    expiration_date: string | null
  }
) {
  const changes: string[] = []

  if (current.name !== next.name) {
    changes.push(
      `Nome: "${current.name}" → "${next.name}"`
    )
  }

  if ((current.supplier ?? '') !== next.supplier) {
    changes.push(
      `Fornitore: "${current.supplier ?? '-'}" → "${next.supplier}"`
    )
  }

  if ((current.category ?? '') !== next.category) {
    changes.push(
      `Tipologia: "${current.category ?? '-'}" → "${next.category}"`
    )
  }

  addNumberChange(
    changes,
    'Prezzo normale',
    current.normal_price,
    next.normal_price
  )

  addNumberChange(
    changes,
    'Prezzo confidenziale',
    current.confidential_price,
    next.confidential_price
  )

  addNumberChange(
    changes,
    'Prezzo SOMS',
    current.sale_price,
    next.sale_price
  )

  if (current.availability !== next.availability) {
    changes.push(
      `Disponibilità: ${current.availability} → ${next.availability}`
    )
  }

  if (current.unit_type !== next.unit_type) {
    changes.push(
      `Unità: ${current.unit_type} → ${next.unit_type}`
    )
  }

  if (
    (current.expiration_date ?? '') !==
    (next.expiration_date ?? '')
  ) {
    changes.push(
      `Scadenza: ${current.expiration_date ?? '-'} → ${next.expiration_date ?? '-'}`
    )
  }

  return changes
}

function addNumberChange(
  changes: string[],
  label: string,
  current: number | null,
  next: number | null
) {
  if (!sameNumber(current, next)) {
    changes.push(
      `${label}: ${formatNumber(current)} → ${formatNumber(next)}`
    )
  }
}

function parseExcelNumber(
  value: unknown,
  required = false
): number | null {
  if (
    value === null ||
    value === undefined ||
    value === ''
  ) {
    if (required) {
      throw new Error(
        'Valore numerico mancante'
      )
    }

    return null
  }

  const parsed =
    typeof value === 'number'
      ? value
      : Number(
          String(value)
            .trim()
            .replace(',', '.')
        )

  if (Number.isNaN(parsed)) {
    throw new Error(
      `Numero non valido: ${String(value)}`
    )
  }

  return parsed
}

function normalizeUnit(value: unknown) {
  const unit = normalizeText(value).toUpperCase()

  if (
    ['PEZZO', 'PEZZI', 'PZ'].includes(unit)
  ) {
    return 'PEZZO'
  }

  if (
    ['KG', 'KILO', 'CHILO', 'CHILI'].includes(unit)
  ) {
    return 'KG'
  }

  if (
    ['LITRO', 'LITRI', 'L'].includes(unit)
  ) {
    return 'LITRO'
  }

  throw new Error(
    `Unità non valida: ${String(value)}`
  )
}

function parseExcelDate(
  value: unknown
): string | null {
  if (
    value === null ||
    value === undefined ||
    value === ''
  ) {
    return null
  }

  if (value instanceof Date) {
    const year = value.getFullYear()
    const month = String(
      value.getMonth() + 1
    ).padStart(2, '0')

    const day = String(
      value.getDate()
    ).padStart(2, '0')

    return `${year}-${month}-${day}`
  }

  const text = String(value).trim()

  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
    return text
  }

  const italian = text.match(
    /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/
  )

  if (italian) {
    const day = italian[1].padStart(2, '0')
    const month = italian[2].padStart(2, '0')

    return `${italian[3]}-${month}-${day}`
  }

  throw new Error(
    `Data non valida: ${text}`
  )
}

function normalizeText(value: unknown) {
  return String(value ?? '')
    .trim()
    .toLowerCase()
}

function sameNumber(
  left: number | null,
  right: number | null
) {
  if (left === null && right === null) {
    return true
  }

  if (left === null || right === null) {
    return false
  }

  return Math.abs(left - right) < 0.000001
}

function formatNumber(
  value: number | null
) {
  if (value === null) {
    return '-'
  }

  return value.toLocaleString('it-IT', {
    maximumFractionDigits: 3,
  })
}

function getProductName(
  products:
    | { name: string }
    | { name: string }[]
    | null
) {
  if (!products) {
    return '-'
  }

  if (Array.isArray(products)) {
    return products[0]?.name ?? '-'
  }

  return products.name
}