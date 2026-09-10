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

  const worksheet = utils.json_to_sheet(rows)

  worksheet['!cols'] = [
    { wch: 38 },
    { wch: 35 },
    { wch: 28 },
    { wch: 22 },
    { wch: 20 },
    { wch: 24 },
    { wch: 18 },
    { wch: 16 },
    { wch: 14 },
    { wch: 16 },
  ]

  const workbook = utils.book_new()
  utils.book_append_sheet(workbook, worksheet, 'Inventario')

  const filename = makeDateTimeFileName('Inventario')
  writeFileXLSX(workbook, filename, { compression: true })

  return filename
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

  const worksheet = utils.json_to_sheet(rows)

  worksheet['!cols'] = [
    { wch: 22 },
    { wch: 35 },
    { wch: 25 },
    { wch: 14 },
    { wch: 14 },
    { wch: 14 },
    { wch: 20 },
    { wch: 45 },
  ]

  const workbook = utils.book_new()
  utils.book_append_sheet(workbook, worksheet, 'Movimenti')

  const filename = makeDateTimeFileName('Movimenti')
  writeFileXLSX(workbook, filename, { compression: true })

  return filename
}

export function exportSalesExcel(sales: Sale[]) {
  const rows: Record<string, string | number>[] = []

  for (const sale of sales) {
    if (!sale.sale_items || sale.sale_items.length === 0) {
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

  const worksheet = utils.json_to_sheet(rows)

  worksheet['!cols'] = [
    { wch: 38 },
    { wch: 22 },
    { wch: 16 },
    { wch: 16 },
    { wch: 35 },
    { wch: 12 },
    { wch: 20 },
    { wch: 18 },
    { wch: 20 },
    { wch: 40 },
  ]

  const workbook = utils.book_new()
  utils.book_append_sheet(workbook, worksheet, 'Storico vendite')

  const filename = makeDateTimeFileName('Storico_Vendite')
  writeFileXLSX(workbook, filename, { compression: true })

  return filename
}

export async function parseInventoryExcel(
  file: File,
  currentProducts: InventoryRow[]
): Promise<ExcelImportRow[]> {
  const arrayBuffer = await file.arrayBuffer()

  const workbook = read(arrayBuffer, {
    cellDates: true,
  })

  if (workbook.SheetNames.length === 0) {
    throw new Error('Il file Excel non contiene fogli')
  }

  const worksheet = workbook.Sheets[workbook.SheetNames[0]]

  const matrix = utils.sheet_to_json<unknown[]>(worksheet, {
    header: 1,
    defval: '',
    raw: true,
  })

  if (matrix.length === 0) {
    throw new Error('Il file Excel è vuoto')
  }

  const headers = (matrix[0] ?? []).map((value) =>
    String(value).trim()
  )

  const missingHeaders = EXCEL_HEADERS.filter(
    (header) => !headers.includes(header)
  )

  if (missingHeaders.length > 0) {
    throw new Error(
      `Formato Excel non valido. Colonne mancanti: ${missingHeaders.join(', ')}`
    )
  }

  const objects = utils.sheet_to_json<Record<string, unknown>>(worksheet, {
    defval: '',
    raw: true,
  })

  const parsedRows: ExcelImportRow[] = []

  objects.forEach((row, index) => {
    const excelRowNumber = index + 2

    const productIdValue = String(
      row['ID prodotto'] ?? ''
    ).trim()

    const name = String(row['Prodotto'] ?? '').trim()

    if (name === '' && productIdValue === '') {
      return
    }

    if (name === '') {
      throw new Error(`Riga ${excelRowNumber}: prodotto mancante`)
    }

    const supplier = String(row['Fornitore'] ?? '').trim()
    if (!supplier) {
      throw new Error(`Riga ${excelRowNumber}: fornitore mancante`)
    }

    const category = String(row['Tipologia'] ?? '').trim()
    if (!category) {
      throw new Error(`Riga ${excelRowNumber}: tipologia mancante`)
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
          product.name.trim().toLowerCase() === name.toLowerCase() &&
          (product.supplier ?? '').trim().toLowerCase() ===
            supplier.toLowerCase()
      )
    }

    const normalPrice = parseExcelNumber(
      row['Prezzo normale (€)']
    )
    const confidentialPrice = parseExcelNumber(
      row['Prezzo confidenziale (€)']
    )
    const salePrice = parseExcelNumber(row['Prezzo SOMS (€)'])

    const currentAvailability = existing
      ? existing.availability
      : null

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
      product_id: productIdValue || existing?.product_id || null,
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

  if (parsedRows.length === 0) {
    throw new Error('Il file non contiene prodotti da importare')
  }

  const ids = parsedRows
    .map((row) => row.product_id)
    .filter((id): id is string => Boolean(id))

  if (new Set(ids).size !== ids.length) {
    throw new Error('Il file contiene prodotti duplicati')
  }

  return parsedRows
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
    changes.push(`Nome: "${current.name}" → "${next.name}"`)
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

  if (!sameNumber(current.normal_price, next.normal_price)) {
    changes.push(
      `Prezzo normale: ${formatNumber(current.normal_price)} → ${formatNumber(next.normal_price)}`
    )
  }

  if (
    !sameNumber(
      current.confidential_price,
      next.confidential_price
    )
  ) {
    changes.push(
      `Prezzo confidenziale: ${formatNumber(current.confidential_price)} → ${formatNumber(next.confidential_price)}`
    )
  }

  if (!sameNumber(current.sale_price, next.sale_price)) {
    changes.push(
      `Prezzo SOMS: ${formatNumber(current.sale_price)} → ${formatNumber(next.sale_price)}`
    )
  }

  if (current.availability !== next.availability) {
    changes.push(
      `Disponibilità: ${current.availability} → ${next.availability}`
    )
  }

  if (current.unit_type !== next.unit_type) {
    changes.push(`Unità: ${current.unit_type} → ${next.unit_type}`)
  }

  if ((current.expiration_date ?? '') !== (next.expiration_date ?? '')) {
    changes.push(
      `Scadenza: ${current.expiration_date ?? '-'} → ${next.expiration_date ?? '-'}`
    )
  }

  return changes
}

function parseExcelNumber(
  value: unknown,
  required = false
): number | null {
  if (value === null || value === undefined || value === '') {
    if (required) {
      throw new Error('Valore numerico mancante')
    }

    return null
  }

  if (typeof value === 'number') {
    if (Number.isNaN(value)) {
      throw new Error('Numero non valido')
    }

    return value
  }

  const parsed = Number(String(value).trim().replace(',', '.'))

  if (Number.isNaN(parsed)) {
    throw new Error(`Numero non valido: ${String(value)}`)
  }

  return parsed
}

function normalizeUnit(value: unknown) {
  const normalized = String(value ?? '').trim().toUpperCase()

  if (['PEZZO', 'PEZZI', 'PZ'].includes(normalized)) {
    return 'PEZZO'
  }

  if (['KG', 'KILO', 'CHILO', 'CHILI'].includes(normalized)) {
    return 'KG'
  }

  if (['LITRO', 'LITRI', 'L'].includes(normalized)) {
    return 'LITRO'
  }

  throw new Error(`Unità non valida: ${String(value)}`)
}

function parseExcelDate(value: unknown): string | null {
  if (value === null || value === undefined || value === '') {
    return null
  }

  if (value instanceof Date) {
    const year = value.getFullYear()
    const month = String(value.getMonth() + 1).padStart(2, '0')
    const day = String(value.getDate()).padStart(2, '0')
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

  throw new Error(`Data non valida: ${text}`)
}

function sameNumber(
  left: number | null,
  right: number | null
) {
  if (left === null && right === null) return true
  if (left === null || right === null) return false

  return Math.abs(Number(left) - Number(right)) < 0.000001
}

function formatNumber(value: number | null) {
  if (value === null) return '-'
  return Number(value).toLocaleString('it-IT', {
    maximumFractionDigits: 3,
  })
}

function getProductName(
  products:
    | { name: string }
    | { name: string }[]
    | null
) {
  if (!products) return '-'
  if (Array.isArray(products)) return products[0]?.name ?? '-'
  return products.name
}
