export type Page =
  | 'inventory'
  | 'sales'
  | 'load'
  | 'order'
  | 'statistics'
  | 'salesHistory'
  | 'movements'
  | 'inactive'

export type ProductRelation =
  | { name: string }
  | { name: string }[]
  | null

export type InventoryRow = {
  product_id: string
  name: string
  supplier: string | null
  category: string | null
  normal_price: number | null
  confidential_price: number | null
  sale_price: number | null
  availability: number
  unit_type: string
  expiration_date: string | null
}

export type CartItem = {
  product_id: string
  name: string
  quantity: number
  sale_price: number
}

export type OrderItem = {
  product_id: string
  quantity: string
}

export type Sale = {
  id: string
  total_amount: number
  payment_method: string | null
  status: string
  created_at: string
  cancellation_reason: string | null
  sale_items: {
    id: string
    quantity: number
    unit_sale_price: number
    confidential_price_at_sale: number | null
    line_total: number
    products: ProductRelation
  }[]
}

export type Movement = {
  id: string
  movement_type: string
  quantity: number
  stock_before: number
  stock_after: number
  unit_cost: number | null
  notes: string | null
  created_at: string
  products: ProductRelation
}

export type CashflowPoint = {
  date: string
  value: number
}

export type ImportMode = 'update' | 'replace'

export type ExcelImportRow = {
  product_id: string | null
  name: string
  supplier: string
  category: string
  normal_price: number | null
  confidential_price: number | null
  sale_price: number | null
  availability: number
  unit_type: string
  expiration_date: string | null
  currentAvailability: number | null
  difference: number | null
  action: 'NUOVO' | 'AGGIORNA'
  changes: string[]
}
