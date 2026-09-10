import {
  useEffect,
  useMemo,
  useState,
} from 'react'

import {
  read,
  utils,
  writeFileXLSX,
} from 'xlsx'

import { supabase } from './lib/supabase'
import './App.css'

type Page =
  | 'inventory'
  | 'sales'
  | 'load'
  | 'order'
  | 'statistics'
  | 'salesHistory'
  | 'movements'
  | 'inactive'

type ProductRelation =
  | { name: string }
  | { name: string }[]
  | null

type InventoryRow = {
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

type CartItem = {
  product_id: string
  name: string
  quantity: number
  sale_price: number
}

type OrderItem = {
  product_id: string
  quantity: string
}

type Sale = {
  id: string
  total_amount: number
  payment_method: string | null
  status: string
  created_at: string
  cancellation_reason: string | null
  sale_items: any[]
}

type Movement = {
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

type CashflowPoint = {
  date: string
  value: number
}

type ExcelImportRow = {
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
}

const EXCEL_HEADERS = [
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

function App() {
  const [page, setPage] =
    useState<Page>('inventory')

  const [inventory, setInventory] =
    useState<InventoryRow[]>([])

  const [
    inactiveProducts,
    setInactiveProducts,
  ] = useState<InventoryRow[]>([])

  const [
    selectedProduct,
    setSelectedProduct,
  ] =
    useState<InventoryRow | null>(
      null
    )

  const [cart, setCart] =
    useState<CartItem[]>([])

  const [
    orderItems,
    setOrderItems,
  ] =
    useState<OrderItem[]>([])

  const [sales, setSales] =
    useState<Sale[]>([])

  const [movements, setMovements] =
    useState<Movement[]>([])

  const [
    excelImportRows,
    setExcelImportRows,
  ] =
    useState<ExcelImportRow[]>([])

  const [
    excelFileName,
    setExcelFileName,
  ] = useState('')

  const [
    importingExcel,
    setImportingExcel,
  ] = useState(false)

  const [
    userEmail,
    setUserEmail,
  ] =
    useState<string | null>(null)

  const [email, setEmail] =
    useState('')

  const [password, setPassword] =
    useState('')

  const [
    paymentMethod,
    setPaymentMethod,
  ] = useState('CONTANTI')

  const [
    loadProductId,
    setLoadProductId,
  ] = useState('')

  const [
    loadQuantity,
    setLoadQuantity,
  ] = useState('')

  const [
    loadNotes,
    setLoadNotes,
  ] = useState('')

  const [
    showNewProduct,
    setShowNewProduct,
  ] = useState(false)

  const [newName, setNewName] =
    useState('')

  const [
    newSupplier,
    setNewSupplier,
  ] = useState('')

  const [
    newCategory,
    setNewCategory,
  ] = useState('')

  const [
    newNormalPrice,
    setNewNormalPrice,
  ] = useState('')

  const [
    newConfidentialPrice,
    setNewConfidentialPrice,
  ] = useState('')

  const [
    newSalePrice,
    setNewSalePrice,
  ] = useState('')

  const [
    newUnitType,
    setNewUnitType,
  ] = useState('PEZZO')

  const [
    newExpirationDate,
    setNewExpirationDate,
  ] = useState('')

  const [
    errorMessage,
    setErrorMessage,
  ] =
    useState<string | null>(null)

  const [
    successMessage,
    setSuccessMessage,
  ] =
    useState<string | null>(null)

  function clearMessages() {
    setErrorMessage(null)
    setSuccessMessage(null)
  }

  function parseDecimal(
    value: string
  ): number | null {
    const cleaned =
      value.trim().replace(',', '.')

    if (cleaned === '') {
      return null
    }

    const number =
      Number(cleaned)

    if (Number.isNaN(number)) {
      return null
    }

    return number
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

    if (
      typeof value === 'number'
    ) {
      if (
        Number.isNaN(value)
      ) {
        throw new Error(
          'Numero non valido'
        )
      }

      return value
    }

    const parsed =
      Number(
        String(value)
          .trim()
          .replace(',', '.')
      )

    if (
      Number.isNaN(parsed)
    ) {
      throw new Error(
        `Numero non valido: ${String(
          value
        )}`
      )
    }

    return parsed
  }

  function normalizeUnit(
    value: unknown
  ) {
    const normalized =
      String(value ?? '')
        .trim()
        .toUpperCase()

    if (
      [
        'PEZZO',
        'PEZZI',
        'PZ',
      ].includes(normalized)
    ) {
      return 'PEZZO'
    }

    if (
      [
        'KG',
        'KILO',
        'CHILO',
        'CHILI',
      ].includes(normalized)
    ) {
      return 'KG'
    }

    if (
      [
        'LITRO',
        'LITRI',
        'L',
      ].includes(normalized)
    ) {
      return 'LITRO'
    }

    throw new Error(
      `Unità non valida: ${String(
        value
      )}`
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

    if (
      value instanceof Date
    ) {
      const year =
        value.getFullYear()

      const month =
        String(
          value.getMonth() + 1
        ).padStart(2, '0')

      const day =
        String(
          value.getDate()
        ).padStart(2, '0')

      return `${year}-${month}-${day}`
    }

    const stringValue =
      String(value).trim()

    if (
      /^\d{4}-\d{2}-\d{2}$/.test(
        stringValue
      )
    ) {
      return stringValue
    }

    const italian =
      stringValue.match(
        /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/
      )

    if (italian) {
      const day =
        italian[1].padStart(
          2,
          '0'
        )

      const month =
        italian[2].padStart(
          2,
          '0'
        )

      return `${italian[3]}-${month}-${day}`
    }

    throw new Error(
      `Data non valida: ${stringValue}`
    )
  }

  async function loadInventory() {
    const {
      data: products,
      error: productsError,
    } = await supabase
      .from('products')
      .select(`
        id,
        name,
        normal_price,
        confidential_price,
        sale_price,
        unit_type,
        expiration_date,
        suppliers(name),
        categories(name)
      `)
      .eq('active', true)
      .order('name')

    if (productsError) {
      setErrorMessage(
        productsError.message
      )
      return
    }

    const {
      data: stocks,
      error: stocksError,
    } = await supabase
      .from('current_stock')
      .select(
        'product_id, availability'
      )

    if (stocksError) {
      setErrorMessage(
        stocksError.message
      )
      return
    }

    const stockMap =
      new Map(
        (stocks ?? []).map(
          (stock) => [
            stock.product_id,
            Number(
              stock.availability
            ),
          ]
        )
      )

    const rows: InventoryRow[] =
      (products ?? []).map(
        (item: any) => ({
          product_id: item.id,
          name: item.name,
          supplier:
            getProductName(
              item.suppliers
            ),
          category:
            getProductName(
              item.categories
            ),
          normal_price:
            item.normal_price,
          confidential_price:
            item.confidential_price,
          sale_price:
            item.sale_price,
          availability:
            stockMap.get(
              item.id
            ) ?? 0,
          unit_type:
            item.unit_type ??
            'PEZZO',
          expiration_date:
            item.expiration_date ??
            null,
        })
      )

    setInventory(rows)
  }

  async function loadInactiveProducts() {
    const {
      data: products,
      error: productsError,
    } = await supabase
      .from('products')
      .select(`
        id,
        name,
        normal_price,
        confidential_price,
        sale_price,
        unit_type,
        expiration_date,
        suppliers(name),
        categories(name)
      `)
      .eq('active', false)
      .order('name')

    if (productsError) {
      setErrorMessage(
        productsError.message
      )
      return
    }

    const {
      data: stocks,
      error: stocksError,
    } = await supabase
      .from('current_stock')
      .select(
        'product_id, availability'
      )

    if (stocksError) {
      setErrorMessage(
        stocksError.message
      )
      return
    }

    const stockMap =
      new Map(
        (stocks ?? []).map(
          (stock) => [
            stock.product_id,
            Number(
              stock.availability
            ),
          ]
        )
      )

    const rows: InventoryRow[] =
      (products ?? []).map(
        (item: any) => ({
          product_id: item.id,
          name: item.name,
          supplier:
            getProductName(
              item.suppliers
            ),
          category:
            getProductName(
              item.categories
            ),
          normal_price:
            item.normal_price,
          confidential_price:
            item.confidential_price,
          sale_price:
            item.sale_price,
          availability:
            stockMap.get(
              item.id
            ) ?? 0,
          unit_type:
            item.unit_type ??
            'PEZZO',
          expiration_date:
            item.expiration_date ??
            null,
        })
      )

    setInactiveProducts(rows)
  }

  async function loadSales() {
    const { data, error } =
      await supabase
        .from('sales')
        .select(`
          id,
          total_amount,
          payment_method,
          status,
          created_at,
          cancellation_reason,
          sale_items(
            id,
            quantity,
            unit_sale_price,
            confidential_price_at_sale,
            line_total,
            products(name)
          )
        `)
        .order(
          'created_at',
          {
            ascending: false,
          }
        )

    if (error) {
      setErrorMessage(
        error.message
      )
      return
    }

    setSales(
      (data ?? []) as Sale[]
    )
  }

  async function loadMovements() {
    const { data, error } =
      await supabase
        .from(
          'stock_movements'
        )
        .select(`
          id,
          movement_type,
          quantity,
          stock_before,
          stock_after,
          unit_cost,
          notes,
          created_at,
          products(name)
        `)
        .order(
          'created_at',
          {
            ascending: false,
          }
        )

    if (error) {
      setErrorMessage(
        error.message
      )
      return
    }

    const normalized:
      Movement[] =
        (data ?? []).map(
          (item: any) => ({
            id: item.id,
            movement_type:
              item.movement_type,
            quantity:
              Number(
                item.quantity
              ),
            stock_before:
              Number(
                item.stock_before
              ),
            stock_after:
              Number(
                item.stock_after
              ),
            unit_cost:
              item.unit_cost ===
              null
                ? null
                : Number(
                    item.unit_cost
                  ),
            notes:
              item.notes,
            created_at:
              item.created_at,
            products:
              item.products ??
              null,
          })
        )

    setMovements(
      normalized
    )
  }

  async function refreshAll() {
    await loadInventory()
    await loadInactiveProducts()
    await loadSales()
    await loadMovements()
  }

  async function login() {
    clearMessages()

    const { error } =
      await supabase.auth
        .signInWithPassword({
          email,
          password,
        })

    if (error) {
      setErrorMessage(
        error.message
      )
    }
  }

  async function logout() {
    await supabase.auth.signOut()
  }

  async function updateExpirationDate(
    productId: string,
    value: string
  ) {
    clearMessages()

    setInventory(
      (current) =>
        current.map(
          (item) =>
            item.product_id ===
            productId
              ? {
                  ...item,
                  expiration_date:
                    value ||
                    null,
                }
              : item
        )
    )

    setSelectedProduct(
      (current) =>
        current?.product_id ===
        productId
          ? {
              ...current,
              expiration_date:
                value || null,
            }
          : current
    )

    const { error } =
      await supabase
        .from('products')
        .update({
          expiration_date:
            value || null,
        })
        .eq(
          'id',
          productId
        )

    if (error) {
      setErrorMessage(
        error.message
      )

      await loadInventory()
      return
    }

    setSuccessMessage(
      'Data di scadenza aggiornata'
    )
  }

  async function deactivateProduct(
    product: InventoryRow
  ) {
    clearMessages()

    const confirmed =
      window.confirm(
        `Vuoi disattivare "${product.name}"?\n\n` +
          `Il prodotto non sarà più disponibile nelle operazioni correnti.\n` +
          `Vendite e movimenti storici resteranno conservati.`
      )

    if (!confirmed) {
      return
    }

    const { error } =
      await supabase
        .from('products')
        .update({
          active: false,
        })
        .eq(
          'id',
          product.product_id
        )

    if (error) {
      setErrorMessage(
        error.message
      )
      return
    }

    setSelectedProduct(
      null
    )

    await refreshAll()

    setSuccessMessage(
      `${product.name} è stato disattivato`
    )
  }

  async function reactivateProduct(
    product: InventoryRow
  ) {
    clearMessages()

    const { error } =
      await supabase
        .from('products')
        .update({
          active: true,
        })
        .eq(
          'id',
          product.product_id
        )

    if (error) {
      setErrorMessage(
        error.message
      )
      return
    }

    await refreshAll()

    setSuccessMessage(
      `${product.name} è stato riattivato`
    )
  }

  function addToCart(
    product: InventoryRow
  ) {
    clearMessages()

    if (
      product.sale_price ===
      null
    ) {
      setErrorMessage(
        'Questo prodotto non ha un prezzo di vendita'
      )
      return
    }

    if (
      product.availability <=
      0
    ) {
      setErrorMessage(
        'Prodotto non disponibile'
      )
      return
    }

    setCart((current) => {
      const existing =
        current.find(
          (item) =>
            item.product_id ===
            product.product_id
        )

      if (existing) {
        if (
          existing.quantity +
            1 >
          product.availability
        ) {
          setErrorMessage(
            'Quantità superiore alla disponibilità'
          )
          return current
        }

        return current.map(
          (item) =>
            item.product_id ===
            product.product_id
              ? {
                  ...item,
                  quantity:
                    item.quantity +
                    1,
                }
              : item
        )
      }

      return [
        ...current,
        {
          product_id:
            product.product_id,
          name:
            product.name,
          quantity: 1,
          sale_price:
            Number(
              product.sale_price
            ),
        },
      ]
    })
  }

  function removeFromCart(
    productId: string
  ) {
    setCart((current) =>
      current.filter(
        (item) =>
          item.product_id !==
          productId
      )
    )
  }

  function changeQuantity(
    productId: string,
    quantity: number
  ) {
    const product =
      inventory.find(
        (item) =>
          item.product_id ===
          productId
      )

    if (!product) {
      return
    }

    if (quantity <= 0) {
      removeFromCart(
        productId
      )
      return
    }

    if (
      quantity >
      product.availability
    ) {
      setErrorMessage(
        'Quantità superiore alla disponibilità'
      )
      return
    }

    setCart((current) =>
      current.map(
        (item) =>
          item.product_id ===
          productId
            ? {
                ...item,
                quantity,
              }
            : item
      )
    )
  }

  async function registerSale() {
    clearMessages()

    if (
      cart.length === 0
    ) {
      setErrorMessage(
        'Aggiungi almeno un prodotto'
      )
      return
    }

    const items =
      cart.map(
        (item) => ({
          product_id:
            item.product_id,
          quantity:
            item.quantity,
        })
      )

    const { error } =
      await supabase.rpc(
        'register_sale',
        {
          p_items:
            items,
          p_payment_method:
            paymentMethod,
          p_notes: null,
        }
      )

    if (error) {
      setErrorMessage(
        error.message
      )
      return
    }

    setCart([])

    setSuccessMessage(
      'Vendita registrata correttamente'
    )

    await refreshAll()
  }

  async function cancelSale(
    saleId: string
  ) {
    clearMessages()

    const reason =
      window.prompt(
        'Inserisci il motivo dell’annullamento:'
      )

    if (
      !reason ||
      reason.trim() === ''
    ) {
      return
    }

    const confirmed =
      window.confirm(
        'Confermi l’annullamento? La merce verrà restituita al magazzino.'
      )

    if (!confirmed) {
      return
    }

    const { error } =
      await supabase.rpc(
        'cancel_sale',
        {
          p_sale_id:
            saleId,
          p_reason:
            reason.trim(),
        }
      )

    if (error) {
      setErrorMessage(
        error.message
      )
      return
    }

    setSuccessMessage(
      'Vendita annullata correttamente'
    )

    await refreshAll()
  }

  async function registerLoad() {
    clearMessages()

    if (
      !loadProductId
    ) {
      setErrorMessage(
        'Seleziona un prodotto'
      )
      return
    }

    const quantity =
      parseDecimal(
        loadQuantity
      )

    if (
      quantity === null ||
      quantity <= 0
    ) {
      setErrorMessage(
        'Inserisci una quantità valida maggiore di zero'
      )
      return
    }

    const { error } =
      await supabase.rpc(
        'register_stock_load',
        {
          p_product_id:
            loadProductId,
          p_quantity:
            quantity,
          p_notes:
            loadNotes.trim() ||
            null,
        }
      )

    if (error) {
      setErrorMessage(
        error.message
      )
      return
    }

    setLoadQuantity('')
    setLoadNotes('')

    setSuccessMessage(
      'Carico registrato correttamente'
    )

    await refreshAll()
  }

  async function createNewProduct() {
    clearMessages()

    if (
      !newName.trim()
    ) {
      setErrorMessage(
        'Inserisci il nome del prodotto'
      )
      return
    }

    if (
      !newSupplier.trim()
    ) {
      setErrorMessage(
        'Inserisci il fornitore'
      )
      return
    }

    if (
      !newCategory.trim()
    ) {
      setErrorMessage(
        'Inserisci la categoria'
      )
      return
    }

    const normalPrice =
      parseDecimal(
        newNormalPrice
      )

    const confidentialPrice =
      parseDecimal(
        newConfidentialPrice
      )

    const salePrice =
      parseDecimal(
        newSalePrice
      )

    let supplierId:
      string | null =
        null

    const {
      data:
        supplierExisting,
      error:
        supplierSearchError,
    } = await supabase
      .from('suppliers')
      .select('id')
      .eq(
        'name',
        newSupplier.trim()
      )
      .maybeSingle()

    if (
      supplierSearchError
    ) {
      setErrorMessage(
        supplierSearchError.message
      )
      return
    }

    if (
      supplierExisting
    ) {
      supplierId =
        supplierExisting.id
    } else {
      const {
        data:
          createdSupplier,
        error,
      } = await supabase
        .from('suppliers')
        .insert({
          name:
            newSupplier.trim(),
        })
        .select('id')
        .single()

      if (error) {
        setErrorMessage(
          error.message
        )
        return
      }

      supplierId =
        createdSupplier.id
    }

    let categoryId:
      string | null =
        null

    const {
      data:
        categoryExisting,
      error:
        categorySearchError,
    } = await supabase
      .from('categories')
      .select('id')
      .eq(
        'name',
        newCategory.trim()
      )
      .maybeSingle()

    if (
      categorySearchError
    ) {
      setErrorMessage(
        categorySearchError.message
      )
      return
    }

    if (
      categoryExisting
    ) {
      categoryId =
        categoryExisting.id
    } else {
      const {
        data:
          createdCategory,
        error,
      } = await supabase
        .from('categories')
        .insert({
          name:
            newCategory.trim(),
        })
        .select('id')
        .single()

      if (error) {
        setErrorMessage(
          error.message
        )
        return
      }

      categoryId =
        createdCategory.id
    }

    const {
      data: product,
      error:
        productError,
    } = await supabase
      .from('products')
      .insert({
        name:
          newName.trim(),
        supplier_id:
          supplierId,
        category_id:
          categoryId,
        normal_price:
          normalPrice,
        confidential_price:
          confidentialPrice,
        sale_price:
          salePrice,
        unit_type:
          newUnitType,
        expiration_date:
          newExpirationDate ||
          null,
        active: true,
      })
      .select('id')
      .single()

    if (
      productError
    ) {
      setErrorMessage(
        productError.message
      )
      return
    }

    await loadInventory()

    setLoadProductId(
      product.id
    )

    setNewName('')
    setNewSupplier('')
    setNewCategory('')
    setNewNormalPrice('')
    setNewConfidentialPrice('')
    setNewSalePrice('')
    setNewUnitType('PEZZO')
    setNewExpirationDate('')
    setShowNewProduct(false)

    setSuccessMessage(
      'Prodotto creato. Ora inserisci la quantità ricevuta.'
    )
  }

  function changeOrderQuantity(
    productId: string,
    value: string
  ) {
    setOrderItems(
      (current) => {
        const existing =
          current.find(
            (item) =>
              item.product_id ===
              productId
          )

        if (existing) {
          return current.map(
            (item) =>
              item.product_id ===
              productId
                ? {
                    ...item,
                    quantity:
                      value,
                  }
                : item
          )
        }

        return [
          ...current,
          {
            product_id:
              productId,
            quantity: value,
          },
        ]
      }
    )
  }

  function getOrderQuantity(
    productId: string
  ) {
    return (
      orderItems.find(
        (item) =>
          item.product_id ===
          productId
      )?.quantity ?? ''
    )
  }

  function getPurchasePrice(
    product: InventoryRow
  ) {
    const confidential =
      Number(
        product.confidential_price ??
          0
      )

    if (
      confidential > 0
    ) {
      return confidential
    }

    return Number(
      product.normal_price ??
        0
    )
  }

  function makeDateTimeFileName(
    prefix: string
  ) {
    const now =
      new Date()

    const date = [
      now.getFullYear(),
      String(
        now.getMonth() + 1
      ).padStart(2, '0'),
      String(
        now.getDate()
      ).padStart(2, '0'),
    ].join('-')

    const time = [
      String(
        now.getHours()
      ).padStart(2, '0'),
      String(
        now.getMinutes()
      ).padStart(2, '0'),
    ].join('-')

    return `${prefix}_${date}_${time}.xlsx`
  }

  function exportInventoryExcel() {
    clearMessages()

    if (
      inventory.length === 0
    ) {
      setErrorMessage(
        'L’inventario è vuoto'
      )
      return
    }

    const rows =
      inventory.map(
        (product) => ({
          Prodotto:
            product.name,
          Fornitore:
            product.supplier ??
            '',
          Tipologia:
            product.category ??
            '',
          'Prezzo normale (€)':
            product.normal_price ??
            '',
          'Prezzo confidenziale (€)':
            product.confidential_price ??
            '',
          'Prezzo SOMS (€)':
            product.sale_price ??
            '',
          Disponibilità:
            product.availability,
          Unità:
            formatUnit(
              product.unit_type
            ),
          Scadenza:
            product.expiration_date
              ? formatDateOnly(
                  product.expiration_date
                )
              : '',
          'Giorni residui':
            getDaysToExpiration(
              product.expiration_date
            ) ?? '',
        })
      )

    const worksheet =
      utils.json_to_sheet(
        rows
      )

    worksheet['!cols'] = [
      { wch: 35 },
      { wch: 28 },
      { wch: 22 },
      { wch: 20 },
      { wch: 24 },
      { wch: 18 },
      { wch: 16 },
      { wch: 12 },
      { wch: 16 },
      { wch: 16 },
    ]

    const workbook =
      utils.book_new()

    utils.book_append_sheet(
      workbook,
      worksheet,
      'Inventario'
    )

    const filename =
      makeDateTimeFileName(
        'Inventario'
      )

    writeFileXLSX(
      workbook,
      filename,
      {
        compression: true,
      }
    )

    setSuccessMessage(
      `File ${filename} generato correttamente`
    )
  }

  function downloadExcelTemplate() {
    clearMessages()

    const rows =
      inventory.map(
        (product) => ({
          'ID prodotto':
            product.product_id,
          Prodotto:
            product.name,
          Fornitore:
            product.supplier ??
            '',
          Tipologia:
            product.category ??
            '',
          'Prezzo normale (€)':
            product.normal_price ??
            '',
          'Prezzo confidenziale (€)':
            product.confidential_price ??
            '',
          'Prezzo SOMS (€)':
            product.sale_price ??
            '',
          Disponibilità:
            product.availability,
          Unità:
            product.unit_type,
          Scadenza:
            product.expiration_date ??
            '',
        })
      )

    /*
     * Aggiungiamo una riga vuota
     * per rendere evidente che
     * è possibile creare nuovi
     * prodotti lasciando ID vuoto.
     */
    rows.push({
      'ID prodotto': '',
      Prodotto: '',
      Fornitore: '',
      Tipologia: '',
      'Prezzo normale (€)': '',
      'Prezzo confidenziale (€)': '',
      'Prezzo SOMS (€)': '',
      Disponibilità: '' as any,
      Unità: '',
      Scadenza: '',
    })

    const worksheet =
      utils.json_to_sheet(
        rows,
        {
          header:
            EXCEL_HEADERS,
        }
      )

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

    const workbook =
      utils.book_new()

    utils.book_append_sheet(
      workbook,
      worksheet,
      'Importazione'
    )

    const filename =
      makeDateTimeFileName(
        'Modello_Inventario'
      )

    writeFileXLSX(
      workbook,
      filename,
      {
        compression: true,
      }
    )
  }

  async function handleExcelUpload(
    file: File
  ) {
    clearMessages()

    try {
      const arrayBuffer =
        await file.arrayBuffer()

      const workbook =
        read(
          arrayBuffer,
          {
            cellDates: true,
          }
        )

      if (
        workbook.SheetNames
          .length === 0
      ) {
        throw new Error(
          'Il file Excel non contiene fogli'
        )
      }

      const worksheet =
        workbook.Sheets[
          workbook.SheetNames[0]
        ]

      const matrix =
        utils.sheet_to_json<
          unknown[]
        >(
          worksheet,
          {
            header: 1,
            defval: '',
            raw: true,
          }
        )

      if (
        matrix.length < 1
      ) {
        throw new Error(
          'Il file Excel è vuoto'
        )
      }

      const headers =
        (matrix[0] ?? [])
          .map(
            (value) =>
              String(
                value
              ).trim()
          )

      const missingHeaders =
        EXCEL_HEADERS.filter(
          (header) =>
            !headers.includes(
              header
            )
        )

      if (
        missingHeaders.length >
        0
      ) {
        throw new Error(
          `Formato Excel non valido. Colonne mancanti: ${missingHeaders.join(
            ', '
          )}`
        )
      }

      const objects =
        utils.sheet_to_json<
          Record<
            string,
            unknown
          >
        >(
          worksheet,
          {
            defval: '',
            raw: true,
          }
        )

      const allProducts = [
        ...inventory,
        ...inactiveProducts,
      ]

      const parsedRows:
        ExcelImportRow[] =
          []

      objects.forEach(
        (
          row,
          index
        ) => {
          const excelRowNumber =
            index + 2

          const name =
            String(
              row[
                'Prodotto'
              ] ?? ''
            ).trim()

          /*
           * Riga completamente vuota:
           * ignorata.
           */
          if (
            name === '' &&
            String(
              row[
                'ID prodotto'
              ] ?? ''
            ).trim() === ''
          ) {
            return
          }

          if (
            name === ''
          ) {
            throw new Error(
              `Riga ${excelRowNumber}: prodotto mancante`
            )
          }

          const supplier =
            String(
              row[
                'Fornitore'
              ] ?? ''
            ).trim()

          if (
            supplier === ''
          ) {
            throw new Error(
              `Riga ${excelRowNumber}: fornitore mancante`
            )
          }

          const category =
            String(
              row[
                'Tipologia'
              ] ?? ''
            ).trim()

          if (
            category === ''
          ) {
            throw new Error(
              `Riga ${excelRowNumber}: tipologia mancante`
            )
          }

          const productIdValue =
            String(
              row[
                'ID prodotto'
              ] ?? ''
            ).trim()

          const availability =
            parseExcelNumber(
              row[
                'Disponibilità'
              ],
              true
            )

          if (
            availability ===
            null ||
            availability < 0
          ) {
            throw new Error(
              `Riga ${excelRowNumber}: disponibilità non valida`
            )
          }

          const unit =
            normalizeUnit(
              row[
                'Unità'
              ]
            )

          const expirationDate =
            parseExcelDate(
              row[
                'Scadenza'
              ]
            )

          let existing:
            InventoryRow |
            undefined

          if (
            productIdValue
          ) {
            existing =
              allProducts.find(
                (product) =>
                  product.product_id ===
                  productIdValue
              )

            if (!existing) {
              throw new Error(
                `Riga ${excelRowNumber}: ID prodotto non trovato nel database`
              )
            }
          } else {
            existing =
              allProducts.find(
                (product) =>
                  product.name
                    .trim()
                    .toLowerCase() ===
                    name.toLowerCase() &&
                  (
                    product.supplier ??
                    ''
                  )
                    .trim()
                    .toLowerCase() ===
                    supplier.toLowerCase()
              )
          }

          const currentAvailability =
            existing
              ? existing.availability
              : null

          const difference =
            currentAvailability ===
            null
              ? availability
              : availability -
                currentAvailability

          parsedRows.push({
            product_id:
              productIdValue ||
              null,

            name,
            supplier,
            category,

            normal_price:
              parseExcelNumber(
                row[
                  'Prezzo normale (€)'
                ]
              ),

            confidential_price:
              parseExcelNumber(
                row[
                  'Prezzo confidenziale (€)'
                ]
              ),

            sale_price:
              parseExcelNumber(
                row[
                  'Prezzo SOMS (€)'
                ]
              ),

            availability,
            unit_type:
              unit,

            expiration_date:
              expirationDate,

            currentAvailability,
            difference,

            action:
              existing
                ? 'AGGIORNA'
                : 'NUOVO',
          })
        }
      )

      if (
        parsedRows.length ===
        0
      ) {
        throw new Error(
          'Il file non contiene prodotti da importare'
        )
      }

      const duplicateIds =
        parsedRows
          .filter(
            (row) =>
              row.product_id
          )
          .map(
            (row) =>
              row.product_id
          )
          .filter(
            (
              id,
              index,
              array
            ) =>
              array.indexOf(
                id
              ) !== index
          )

      if (
        duplicateIds.length >
        0
      ) {
        throw new Error(
          'Il file contiene ID prodotto duplicati'
        )
      }

      setExcelImportRows(
        parsedRows
      )

      setExcelFileName(
        file.name
      )

      setSuccessMessage(
        'Excel letto correttamente. Controlla l’anteprima prima di confermare.'
      )
    } catch (error) {
      setExcelImportRows(
        []
      )

      setExcelFileName('')

      setErrorMessage(
        error instanceof
          Error
          ? error.message
          : 'Errore durante la lettura del file Excel'
      )
    }
  }

  async function confirmExcelImport() {
    clearMessages()

    if (
      excelImportRows.length ===
      0
    ) {
      setErrorMessage(
        'Non ci sono righe da importare'
      )
      return
    }

    const confirmed =
      window.confirm(
        `Confermi l'importazione di ${excelImportRows.length} prodotti?\n\n` +
          `Le variazioni di disponibilità saranno registrate come movimenti di magazzino.`
      )

    if (!confirmed) {
      return
    }

    setImportingExcel(true)

    const payload =
      excelImportRows.map(
        (row) => ({
          product_id:
            row.product_id,
          name:
            row.name,
          supplier:
            row.supplier,
          category:
            row.category,
          normal_price:
            row.normal_price,
          confidential_price:
            row.confidential_price,
          sale_price:
            row.sale_price,
          availability:
            row.availability,
          unit_type:
            row.unit_type,
          expiration_date:
            row.expiration_date,
        })
      )

    const {
      data,
      error,
    } = await supabase.rpc(
      'import_inventory_excel',
      {
        p_rows:
          payload,
      }
    )

    setImportingExcel(false)

    if (error) {
      setErrorMessage(
        error.message
      )
      return
    }

    await refreshAll()

    setExcelImportRows(
      []
    )

    setExcelFileName('')

    const result =
      data as {
        created?: number
        updated?: number
        adjustments?: number
      } | null

    setSuccessMessage(
      `Importazione completata: ${
        result?.created ?? 0
      } nuovi prodotti, ${
        result?.updated ?? 0
      } aggiornati, ${
        result?.adjustments ??
        0
      } rettifiche di magazzino.`
    )
  }

  function cancelExcelImport() {
    setExcelImportRows(
      []
    )

    setExcelFileName('')

    clearMessages()
  }

  function printInventory() {
    clearMessages()
    window.print()
  }

  function exportOrderExcel() {
    clearMessages()

    const selectedRows =
      orderItems
        .map(
          (orderItem) => {
            const product =
              inventory.find(
                (item) =>
                  item.product_id ===
                  orderItem.product_id
              )

            if (!product) {
              return null
            }

            const quantity =
              parseDecimal(
                orderItem.quantity
              )

            if (
              quantity ===
                null ||
              quantity <= 0
            ) {
              return null
            }

            const unitPrice =
              getPurchasePrice(
                product
              )

            return {
              Prodotto:
                product.name,
              Fornitore:
                product.supplier ??
                '-',
              Quantità:
                quantity,
              Unità:
                formatUnit(
                  product.unit_type
                ),
              'Prezzo unitario (€)':
                unitPrice,
              'Spesa prevista (€)':
                Number(
                  (
                    quantity *
                    unitPrice
                  ).toFixed(2)
                ),
            }
          }
        )
        .filter(
          Boolean
        ) as any[]

    if (
      selectedRows.length ===
      0
    ) {
      setErrorMessage(
        'Inserisci almeno una quantità da ordinare'
      )
      return
    }

    const total =
      selectedRows.reduce(
        (sum, row) =>
          sum +
          row[
            'Spesa prevista (€)'
          ],
        0
      )

    const excelRows = [
      ...selectedRows,
      {
        Prodotto: '',
        Fornitore: '',
        Quantità: '',
        Unità: '',
        'Prezzo unitario (€)':
          '',
        'Spesa prevista (€)':
          '',
      },
      {
        Prodotto:
          'TOTALE',
        Fornitore: '',
        Quantità: '',
        Unità: '',
        'Prezzo unitario (€)':
          '',
        'Spesa prevista (€)':
          Number(
            total.toFixed(
              2
            )
          ),
      },
    ]

    const worksheet =
      utils.json_to_sheet(
        excelRows
      )

    worksheet['!cols'] = [
      { wch: 35 },
      { wch: 28 },
      { wch: 12 },
      { wch: 12 },
      { wch: 20 },
      { wch: 20 },
    ]

    const workbook =
      utils.book_new()

    utils.book_append_sheet(
      workbook,
      worksheet,
      'Ordine'
    )

    const filename =
      makeDateTimeFileName(
        'Ordini'
      )

    writeFileXLSX(
      workbook,
      filename,
      {
        compression: true,
      }
    )

    setSuccessMessage(
      `File ${filename} generato correttamente`
    )
  }

  const selectedLoadProduct =
    inventory.find(
      (product) =>
        product.product_id ===
        loadProductId
    )

  const total =
    useMemo(
      () =>
        cart.reduce(
          (
            sum,
            item
          ) =>
            sum +
            item.quantity *
              item.sale_price,
          0
        ),
      [cart]
    )

  const orderTotal =
    useMemo(() => {
      return orderItems.reduce(
        (
          sum,
          orderItem
        ) => {
          const product =
            inventory.find(
              (item) =>
                item.product_id ===
                orderItem.product_id
            )

          if (!product) {
            return sum
          }

          const quantity =
            parseDecimal(
              orderItem.quantity
            )

          if (
            quantity ===
              null ||
            quantity <= 0
          ) {
            return sum
          }

          return (
            sum +
            quantity *
              getPurchasePrice(
                product
              )
          )
        },
        0
      )
    }, [
      orderItems,
      inventory,
    ])

  const categoryCount =
    useMemo(() => {
      return new Set(
        inventory
          .map(
            (item) =>
              item.category
          )
          .filter(Boolean)
      ).size
    }, [inventory])

  const tiedCapital =
    useMemo(() => {
      return [
        ...inventory,
        ...inactiveProducts,
      ].reduce(
        (
          sum,
          product
        ) =>
          sum +
          product.availability *
            getPurchasePrice(
              product
            ),
        0
      )
    }, [
      inventory,
      inactiveProducts,
    ])

  const inventorySaleValue =
    useMemo(() => {
      return [
        ...inventory,
        ...inactiveProducts,
      ].reduce(
        (
          sum,
          product
        ) =>
          sum +
          product.availability *
            Number(
              product.sale_price ??
                0
            ),
        0
      )
    }, [
      inventory,
      inactiveProducts,
    ])

  const potentialMargin =
    inventorySaleValue -
    tiedCapital

  const completedSalesRevenue =
    useMemo(() => {
      return sales
        .filter(
          (sale) =>
            sale.status ===
            'COMPLETATA'
        )
        .reduce(
          (
            sum,
            sale
          ) =>
            sum +
            Number(
              sale.total_amount
            ),
          0
        )
    }, [sales])

  const purchaseOutflow =
    useMemo(() => {
      return movements
        .filter(
          (movement) =>
            [
              'INVENTARIO_INIZIALE',
              'CARICO',
            ].includes(
              movement.movement_type
            )
        )
        .reduce(
          (
            sum,
            movement
          ) =>
            sum +
            movement.quantity *
              Number(
                movement.unit_cost ??
                  0
              ),
          0
        )
    }, [movements])

  const netCashflow =
    completedSalesRevenue -
    purchaseOutflow

  const cashflowPoints =
    useMemo(() => {
      const events: {
        date: string
        amount: number
      }[] = []

      movements
        .filter(
          (movement) =>
            [
              'INVENTARIO_INIZIALE',
              'CARICO',
            ].includes(
              movement.movement_type
            )
        )
        .forEach(
          (movement) => {
            events.push({
              date:
                movement.created_at,
              amount:
                -movement.quantity *
                Number(
                  movement.unit_cost ??
                    0
                ),
            })
          }
        )

      sales
        .filter(
          (sale) =>
            sale.status ===
            'COMPLETATA'
        )
        .forEach(
          (sale) => {
            events.push({
              date:
                sale.created_at,
              amount:
                Number(
                  sale.total_amount
                ),
            })
          }
        )

      events.sort(
        (a, b) =>
          new Date(
            a.date
          ).getTime() -
          new Date(
            b.date
          ).getTime()
      )

      let cumulative = 0

      return events.map(
        (event) => {
          cumulative +=
            event.amount

          return {
            date:
              event.date,
            value:
              cumulative,
          }
        }
      )
    }, [
      movements,
      sales,
    ])

  const excelNewCount =
    excelImportRows.filter(
      (row) =>
        row.action ===
        'NUOVO'
    ).length

  const excelUpdateCount =
    excelImportRows.filter(
      (row) =>
        row.action ===
        'AGGIORNA'
    ).length

  const excelAdjustmentCount =
    excelImportRows.filter(
      (row) =>
        row.difference !==
          null &&
        row.difference !==
          0
    ).length

  useEffect(() => {
    supabase.auth
      .getSession()
      .then(({ data }) => {
        setUserEmail(
          data.session?.user
            .email ?? null
        )
      })

    const { data } =
      supabase.auth
        .onAuthStateChange(
          (
            _event,
            session
          ) => {
            setUserEmail(
              session?.user
                .email ??
                null
            )
          }
        )

    return () => {
      data.subscription
        .unsubscribe()
    }
  }, [])

  useEffect(() => {
    if (userEmail) {
      refreshAll()
    } else {
      setInventory([])
      setInactiveProducts(
        []
      )
      setSales([])
      setMovements([])
      setCart([])
    }
  }, [userEmail])

  if (!userEmail) {
    return (
      <div className="login-page">
        <div className="login-card">
          <img
            src="/logo.png"
            alt="Logo"
            className="login-logo"
          />

          <h1>
            Gramsci Warehouse
            Management
          </h1>

          <p>
            Accedi per gestire
            il magazzino.
          </p>

          <input
            className="form-control"
            type="email"
            placeholder="Email"
            value={email}
            onChange={(e) =>
              setEmail(
                e.target.value
              )
            }
          />

          <input
            className="form-control"
            type="password"
            placeholder="Password"
            value={password}
            onChange={(e) =>
              setPassword(
                e.target.value
              )
            }
          />

          <button
            className="btn btn-primary btn-full"
            onClick={login}
          >
            Accedi
          </button>

          {errorMessage && (
            <div className="alert alert-error">
              {errorMessage}
            </div>
          )}
        </div>
      </div>
    )
  }

  return (
    <div className="app">
      <header className="app-header no-print">
        <div className="header-logo">
          <img
            src="/logo.png"
            alt="Logo"
            className="logo-image"
          />
        </div>

        <div className="header-title">
          <h1>
            Gramsci Warehouse
            Management
          </h1>

          <span>
            Gestione magazzino
          </span>
        </div>

        <div className="user-area">
          <span>
            {userEmail}
          </span>

          <button
            className="btn btn-secondary"
            onClick={logout}
          >
            Esci
          </button>
        </div>
      </header>

      <nav className="app-nav no-print">
        <NavButton
          label="Inventario"
          active={
            page ===
            'inventory'
          }
          onClick={() =>
            setPage(
              'inventory'
            )
          }
        />

        <NavButton
          label="Vendite"
          active={
            page === 'sales'
          }
          onClick={() =>
            setPage('sales')
          }
        />

        <NavButton
          label="Carico merce"
          active={
            page === 'load'
          }
          onClick={() =>
            setPage('load')
          }
        />

        <NavButton
          label="Ordina"
          active={
            page === 'order'
          }
          onClick={() =>
            setPage('order')
          }
        />

        <NavButton
          label="Statistiche"
          active={
            page ===
            'statistics'
          }
          onClick={() =>
            setPage(
              'statistics'
            )
          }
        />

        <NavButton
          label="Storico vendite"
          active={
            page ===
            'salesHistory'
          }
          onClick={() =>
            setPage(
              'salesHistory'
            )
          }
        />

        <NavButton
          label="Movimenti"
          active={
            page ===
            'movements'
          }
          onClick={() =>
            setPage(
              'movements'
            )
          }
        />

        <NavButton
          label="Prodotti disattivati"
          active={
            page ===
            'inactive'
          }
          onClick={() =>
            setPage(
              'inactive'
            )
          }
        />
      </nav>

      <main className="main-content">

        {page === 'inventory' && (
          <div className="inventory-print-area">

            <div className="print-only print-inventory-header">
              <h1>
                Gramsci Warehouse
                Management
              </h1>

              <h2>
                Inventario
              </h2>

              <p>
                Stampato il{' '}
                {new Date().toLocaleString(
                  'it-IT'
                )}
              </p>
            </div>

            <div className="no-print">
              <PageTitle
                title="Inventario"
                subtitle="Situazione attuale del magazzino"
              />

              <div className="inventory-actions">

                <button
                  className="btn btn-primary"
                  onClick={
                    exportInventoryExcel
                  }
                >
                  Esporta Excel
                </button>

                <button
                  className="btn btn-secondary"
                  onClick={
                    printInventory
                  }
                >
                  Stampa inventario
                </button>

                <button
                  className="btn btn-secondary"
                  onClick={
                    downloadExcelTemplate
                  }
                >
                  Scarica modello Excel
                </button>

                <label className="btn btn-upload">
                  Carica Excel

                  <input
                    className="hidden-file-input"
                    type="file"
                    accept=".xlsx,.xls"
                    onChange={async (
                      e
                    ) => {
                      const file =
                        e.target
                          .files?.[0]

                      if (file) {
                        await handleExcelUpload(
                          file
                        )
                      }

                      e.currentTarget.value =
                        ''
                    }}
                  />
                </label>

              </div>

              <div className="summary-grid">
                <SummaryCard
                  label="Prodotti"
                  value={
                    inventory.length
                  }
                />

                <SummaryCard
                  label="Tipologie"
                  value={
                    categoryCount
                  }
                />

                <SummaryCard
                  label="Prodotti esauriti"
                  value={
                    inventory.filter(
                      (item) =>
                        item.availability <=
                        0
                    ).length
                  }
                />
              </div>

              {excelImportRows.length >
                0 && (
                <div className="excel-import-panel">

                  <div className="excel-import-header">
                    <div>
                      <span className="modal-eyebrow">
                        Anteprima
                        importazione
                      </span>

                      <h3>
                        {excelFileName}
                      </h3>
                    </div>

                    <span className="badge badge-warning">
                      Nessuna modifica
                      ancora applicata
                    </span>
                  </div>

                  <div className="import-summary-grid">
                    <SummaryCard
                      label="Righe valide"
                      value={
                        excelImportRows.length
                      }
                    />

                    <SummaryCard
                      label="Nuovi prodotti"
                      value={
                        excelNewCount
                      }
                    />

                    <SummaryCard
                      label="Prodotti aggiornati"
                      value={
                        excelUpdateCount
                      }
                    />

                    <SummaryCard
                      label="Rettifiche giacenza"
                      value={
                        excelAdjustmentCount
                      }
                    />
                  </div>

                  <div className="table-wrapper import-preview-table">
                    <table>
                      <thead>
                        <tr>
                          <th>
                            Azione
                          </th>
                          <th>
                            Prodotto
                          </th>
                          <th>
                            Fornitore
                          </th>
                          <th>
                            Tipologia
                          </th>
                          <th>
                            Attuale
                          </th>
                          <th>
                            Excel
                          </th>
                          <th>
                            Differenza
                          </th>
                          <th>
                            Unità
                          </th>
                          <th>
                            Scadenza
                          </th>
                        </tr>
                      </thead>

                      <tbody>
                        {excelImportRows.map(
                          (
                            row,
                            index
                          ) => (
                            <tr
                              key={`${row.product_id ?? row.name}-${index}`}
                            >
                              <td>
                                <span
                                  className={
                                    row.action ===
                                    'NUOVO'
                                      ? 'badge badge-warning'
                                      : 'badge badge-neutral'
                                  }
                                >
                                  {
                                    row.action
                                  }
                                </span>
                              </td>

                              <td>
                                <strong>
                                  {
                                    row.name
                                  }
                                </strong>
                              </td>

                              <td>
                                {
                                  row.supplier
                                }
                              </td>

                              <td>
                                {
                                  row.category
                                }
                              </td>

                              <td>
                                {row.currentAvailability ===
                                null
                                  ? '-'
                                  : formatQuantity(
                                      row.currentAvailability,
                                      row.unit_type
                                    )}
                              </td>

                              <td>
                                {formatQuantity(
                                  row.availability,
                                  row.unit_type
                                )}
                              </td>

                              <td>
                                <DifferenceBadge
                                  difference={
                                    row.difference
                                  }
                                  unit={
                                    row.unit_type
                                  }
                                />
                              </td>

                              <td>
                                {formatUnit(
                                  row.unit_type
                                )}
                              </td>

                              <td>
                                {row.expiration_date
                                  ? formatDateOnly(
                                      row.expiration_date
                                    )
                                  : '-'}
                              </td>
                            </tr>
                          )
                        )}
                      </tbody>
                    </table>
                  </div>

                  <div className="import-actions">
                    <button
                      className="btn btn-secondary"
                      onClick={
                        cancelExcelImport
                      }
                      disabled={
                        importingExcel
                      }
                    >
                      Annulla
                    </button>

                    <button
                      className="btn btn-primary btn-large"
                      onClick={
                        confirmExcelImport
                      }
                      disabled={
                        importingExcel
                      }
                    >
                      {importingExcel
                        ? 'Importazione...'
                        : 'Conferma aggiornamento'}
                    </button>
                  </div>

                </div>
              )}

            </div>

            <div className="panel inventory-panel">
              <div className="table-wrapper">
                <table className="inventory-table">
                  <thead>
                    <tr>
                      <th>
                        Prodotto
                      </th>
                      <th>
                        Fornitore
                      </th>
                      <th>
                        Tipo
                      </th>
                      <th>
                        Prezzo
                      </th>
                      <th>
                        P. confidenziale
                      </th>
                      <th>
                        P. SOMS
                      </th>
                      <th>
                        Disponibilità
                      </th>
                      <th>
                        Scadenza
                      </th>
                      <th>
                        Giorni residui
                      </th>
                    </tr>
                  </thead>

                  <tbody>
                    {inventory.map(
                      (item) => {
                        const days =
                          getDaysToExpiration(
                            item.expiration_date
                          )

                        return (
                          <tr
                            key={
                              item.product_id
                            }
                          >
                            <td>
                              <button
                                type="button"
                                className="product-name-button no-print"
                                onClick={() =>
                                  setSelectedProduct(
                                    item
                                  )
                                }
                              >
                                {
                                  item.name
                                }
                              </button>

                              <span className="print-only">
                                {
                                  item.name
                                }
                              </span>
                            </td>

                            <td>
                              {item.supplier ??
                                '-'}
                            </td>

                            <td>
                              {item.category ??
                                '-'}
                            </td>

                            <td>
                              {formatPrice(
                                item.normal_price
                              )}
                            </td>

                            <td>
                              {formatPrice(
                                item.confidential_price
                              )}
                            </td>

                            <td>
                              {formatPrice(
                                item.sale_price
                              )}
                            </td>

                            <td>
                              <span
                                className={
                                  item.availability <=
                                  0
                                    ? 'badge badge-danger'
                                    : item.availability <=
                                      5
                                    ? 'badge badge-warning'
                                    : 'badge badge-success'
                                }
                              >
                                {formatQuantity(
                                  item.availability,
                                  item.unit_type
                                )}
                              </span>
                            </td>

                            <td>
                              <input
                                className="date-input no-print"
                                type="date"
                                value={
                                  item.expiration_date ??
                                  ''
                                }
                                onChange={(
                                  e
                                ) =>
                                  updateExpirationDate(
                                    item.product_id,
                                    e.target
                                      .value
                                  )
                                }
                              />

                              <span className="print-only">
                                {item.expiration_date
                                  ? formatDateOnly(
                                      item.expiration_date
                                    )
                                  : '-'}
                              </span>
                            </td>

                            <td>
                              <ExpirationBadge
                                days={
                                  days
                                }
                              />
                            </td>
                          </tr>
                        )
                      }
                    )}
                  </tbody>
                </table>
              </div>
            </div>

          </div>
        )}

        {page === 'sales' && (
          <>
            <PageTitle
              title="Nuova vendita"
              subtitle="Seleziona i prodotti e registra la vendita"
            />

            <div className="two-column-layout">

              <div className="panel">
                <h3>
                  Prodotti
                </h3>

                <div className="product-list">
                  {inventory.map(
                    (product) => (
                      <div
                        className="product-row"
                        key={
                          product.product_id
                        }
                      >
                        <div>
                          <strong>
                            {
                              product.name
                            }
                          </strong>

                          <span>
                            {formatPrice(
                              product.sale_price
                            )}
                            {' · '}
                            {formatQuantity(
                              product.availability,
                              product.unit_type
                            )}
                          </span>
                        </div>

                        <button
                          className="btn btn-primary"
                          onClick={() =>
                            addToCart(
                              product
                            )
                          }
                        >
                          Aggiungi
                        </button>
                      </div>
                    )
                  )}
                </div>
              </div>

              <div className="panel sticky-panel">
                <h3>
                  Carrello
                </h3>

                {cart.length ===
                  0 && (
                  <div className="empty-state">
                    Nessun prodotto
                    nel carrello.
                  </div>
                )}

                {cart.map(
                  (item) => (
                    <div
                      className="cart-row"
                      key={
                        item.product_id
                      }
                    >
                      <div>
                        <strong>
                          {
                            item.name
                          }
                        </strong>

                        <span>
                          {formatPrice(
                            item.quantity *
                              item.sale_price
                          )}
                        </span>
                      </div>

                      <div className="cart-actions">
                        <input
                          className="quantity-input"
                          type="number"
                          min="0.001"
                          step="0.001"
                          value={
                            item.quantity
                          }
                          onChange={(
                            e
                          ) =>
                            changeQuantity(
                              item.product_id,
                              Number(
                                e.target
                                  .value
                              )
                            )
                          }
                        />

                        <button
                          className="btn btn-danger btn-small"
                          onClick={() =>
                            removeFromCart(
                              item.product_id
                            )
                          }
                        >
                          Rimuovi
                        </button>
                      </div>
                    </div>
                  )
                )}

                <div className="checkout-total">
                  <span>
                    Totale
                  </span>

                  <strong>
                    {formatPrice(
                      total
                    )}
                  </strong>
                </div>

                <label>
                  Metodo di
                  pagamento
                </label>

                <select
                  className="form-control"
                  value={
                    paymentMethod
                  }
                  onChange={(e) =>
                    setPaymentMethod(
                      e.target.value
                    )
                  }
                >
                  <option value="CONTANTI">
                    Contanti
                  </option>

                  <option value="CARTA">
                    Carta
                  </option>

                  <option value="ALTRO">
                    Altro
                  </option>
                </select>

                <button
                  className="btn btn-primary btn-full btn-large"
                  onClick={
                    registerSale
                  }
                >
                  Registra vendita
                </button>
              </div>

            </div>
          </>
        )}

        {page === 'load' && (
          <>
            <PageTitle
              title="Carico merce"
              subtitle="Registra merce ricevuta o crea un nuovo prodotto"
            />

            <div className="panel form-panel">

              <button
                className="btn btn-secondary"
                onClick={() => {
                  clearMessages()

                  setShowNewProduct(
                    !showNewProduct
                  )
                }}
              >
                {showNewProduct
                  ? 'Chiudi nuovo prodotto'
                  : '+ Nuovo prodotto'}
              </button>

              {showNewProduct && (
                <div className="nested-panel">
                  <h3>
                    Nuovo prodotto
                  </h3>

                  <div className="form-grid">

                    <FormField label="Nome prodotto">
                      <input
                        className="form-control"
                        value={
                          newName
                        }
                        onChange={(
                          e
                        ) =>
                          setNewName(
                            e.target
                              .value
                          )
                        }
                      />
                    </FormField>

                    <FormField label="Fornitore">
                      <input
                        className="form-control"
                        value={
                          newSupplier
                        }
                        onChange={(
                          e
                        ) =>
                          setNewSupplier(
                            e.target
                              .value
                          )
                        }
                      />
                    </FormField>

                    <FormField label="Categoria">
                      <input
                        className="form-control"
                        value={
                          newCategory
                        }
                        onChange={(
                          e
                        ) =>
                          setNewCategory(
                            e.target
                              .value
                          )
                        }
                      />
                    </FormField>

                    <FormField label="Unità di misura">
                      <select
                        className="form-control"
                        value={
                          newUnitType
                        }
                        onChange={(
                          e
                        ) =>
                          setNewUnitType(
                            e.target
                              .value
                          )
                        }
                      >
                        <option value="PEZZO">
                          Pezzo
                        </option>

                        <option value="KG">
                          Kg
                        </option>

                        <option value="LITRO">
                          Litro
                        </option>
                      </select>
                    </FormField>

                    <FormField label="Prezzo normale">
                      <input
                        className="form-control"
                        value={
                          newNormalPrice
                        }
                        onChange={(
                          e
                        ) =>
                          setNewNormalPrice(
                            e.target
                              .value
                          )
                        }
                      />
                    </FormField>

                    <FormField label="Prezzo confidenziale">
                      <input
                        className="form-control"
                        value={
                          newConfidentialPrice
                        }
                        onChange={(
                          e
                        ) =>
                          setNewConfidentialPrice(
                            e.target
                              .value
                          )
                        }
                      />
                    </FormField>

                    <FormField label="Prezzo SOMS">
                      <input
                        className="form-control"
                        value={
                          newSalePrice
                        }
                        onChange={(
                          e
                        ) =>
                          setNewSalePrice(
                            e.target
                              .value
                          )
                        }
                      />
                    </FormField>

                    <FormField label="Data di scadenza">
                      <input
                        className="form-control"
                        type="date"
                        value={
                          newExpirationDate
                        }
                        onChange={(
                          e
                        ) =>
                          setNewExpirationDate(
                            e.target
                              .value
                          )
                        }
                      />
                    </FormField>

                  </div>

                  <button
                    className="btn btn-primary"
                    onClick={
                      createNewProduct
                    }
                  >
                    Crea prodotto
                  </button>
                </div>
              )}

              <div className="section-divider" />

              <h3>
                Registra carico
              </h3>

              <FormField label="Prodotto">
                <select
                  className="form-control"
                  value={
                    loadProductId
                  }
                  onChange={(e) => {
                    setLoadProductId(
                      e.target.value
                    )

                    setLoadQuantity(
                      ''
                    )

                    clearMessages()
                  }}
                >
                  <option value="">
                    Seleziona prodotto
                  </option>

                  {inventory.map(
                    (product) => (
                      <option
                        key={
                          product.product_id
                        }
                        value={
                          product.product_id
                        }
                      >
                        {
                          product.name
                        }
                      </option>
                    )
                  )}
                </select>
              </FormField>

              {selectedLoadProduct && (
                <div className="product-info-card">
                  <div>
                    <span>
                      Fornitore
                    </span>

                    <strong>
                      {selectedLoadProduct.supplier ??
                        '-'}
                    </strong>
                  </div>

                  <div>
                    <span>
                      Categoria
                    </span>

                    <strong>
                      {selectedLoadProduct.category ??
                        '-'}
                    </strong>
                  </div>

                  <div>
                    <span>
                      Disponibilità
                    </span>

                    <strong>
                      {formatQuantity(
                        selectedLoadProduct.availability,
                        selectedLoadProduct.unit_type
                      )}
                    </strong>
                  </div>
                </div>
              )}

              <FormField
                label={
                  selectedLoadProduct?.unit_type ===
                  'KG'
                    ? 'Peso ricevuto (kg)'
                    : selectedLoadProduct?.unit_type ===
                      'LITRO'
                    ? 'Volume ricevuto (litri)'
                    : 'Quantità ricevuta (pezzi)'
                }
              >
                <input
                  className="form-control"
                  type="text"
                  inputMode="decimal"
                  value={
                    loadQuantity
                  }
                  onChange={(e) =>
                    setLoadQuantity(
                      e.target.value
                    )
                  }
                />
              </FormField>

              <FormField label="Note / DDT / riferimento">
                <textarea
                  className="form-control textarea"
                  value={
                    loadNotes
                  }
                  onChange={(e) =>
                    setLoadNotes(
                      e.target.value
                    )
                  }
                />
              </FormField>

              <button
                className="btn btn-primary"
                onClick={
                  registerLoad
                }
              >
                Registra carico
              </button>

            </div>
          </>
        )}

        {page === 'order' && (
          <>
            <PageTitle
              title="Ordina"
              subtitle="Prepara un ordine e genera il file Excel"
            />

            <div className="panel">
              <div className="table-wrapper">
                <table>
                  <thead>
                    <tr>
                      <th>
                        Prodotto
                      </th>
                      <th>
                        Fornitore
                      </th>
                      <th>
                        Disponibilità
                      </th>
                      <th>
                        Prezzo acquisto
                      </th>
                      <th>
                        Quantità da ordinare
                      </th>
                      <th>
                        Spesa prevista
                      </th>
                    </tr>
                  </thead>

                  <tbody>
                    {inventory.map(
                      (product) => {
                        const quantityText =
                          getOrderQuantity(
                            product.product_id
                          )

                        const quantity =
                          parseDecimal(
                            quantityText
                          )

                        const purchasePrice =
                          getPurchasePrice(
                            product
                          )

                        const rowTotal =
                          quantity !==
                            null &&
                          quantity > 0
                            ? quantity *
                              purchasePrice
                            : 0

                        return (
                          <tr
                            key={
                              product.product_id
                            }
                          >
                            <td>
                              <strong>
                                {
                                  product.name
                                }
                              </strong>
                            </td>

                            <td>
                              {product.supplier ??
                                '-'}
                            </td>

                            <td>
                              {formatQuantity(
                                product.availability,
                                product.unit_type
                              )}
                            </td>

                            <td>
                              {formatPrice(
                                purchasePrice
                              )}
                            </td>

                            <td>
                              <input
                                className="order-input"
                                type="text"
                                inputMode="decimal"
                                value={
                                  quantityText
                                }
                                onChange={(
                                  e
                                ) =>
                                  changeOrderQuantity(
                                    product.product_id,
                                    e.target
                                      .value
                                  )
                                }
                              />
                            </td>

                            <td>
                              {rowTotal >
                              0
                                ? formatPrice(
                                    rowTotal
                                  )
                                : '-'}
                            </td>
                          </tr>
                        )
                      }
                    )}
                  </tbody>
                </table>
              </div>

              <div className="order-footer">
                <div>
                  <span>
                    Spesa totale
                    prevista
                  </span>

                  <strong>
                    {formatPrice(
                      orderTotal
                    )}
                  </strong>
                </div>

                <button
                  className="btn btn-primary"
                  onClick={
                    exportOrderExcel
                  }
                >
                  Esporta Excel
                </button>
              </div>
            </div>
          </>
        )}

        {page === 'statistics' && (
          <>
            <PageTitle
              title="Statistiche"
              subtitle="Valore del magazzino e andamento finanziario"
            />

            <div className="stats-grid">
              <SummaryCard
                label="Capitale fermo"
                value={
                  formatPrice(
                    tiedCapital
                  )
                }
              />

              <SummaryCard
                label="Valore potenziale di vendita"
                value={
                  formatPrice(
                    inventorySaleValue
                  )
                }
              />

              <SummaryCard
                label="Margine potenziale"
                value={
                  formatPrice(
                    potentialMargin
                  )
                }
              />

              <SummaryCard
                label="Ricavi vendite"
                value={
                  formatPrice(
                    completedSalesRevenue
                  )
                }
              />

              <SummaryCard
                label="Capitale investito storico"
                value={
                  formatPrice(
                    purchaseOutflow
                  )
                }
              />

              <SummaryCard
                label="Cashflow cumulato"
                value={
                  formatPrice(
                    netCashflow
                  )
                }
              />
            </div>

            <div className="panel">
              <h3>
                Cashflow cumulato
              </h3>

              <p className="chart-description">
                Parte dal capitale
                investito
                nell'inventario
                iniziale, sottrae i
                successivi carichi
                e aggiunge gli
                incassi delle
                vendite completate.
              </p>

              <CashflowChart
                points={
                  cashflowPoints
                }
              />
            </div>
          </>
        )}

        {page ===
          'salesHistory' && (
          <>
            <PageTitle
              title="Storico vendite"
              subtitle="Vendite registrate e annullamenti"
            />

            <div className="cards-list">
              {sales.map(
                (sale) => (
                  <div
                    key={
                      sale.id
                    }
                    className="sale-card"
                  >
                    <div className="sale-card-header">
                      <div>
                        <strong>
                          {formatDate(
                            sale.created_at
                          )}
                        </strong>

                        <span>
                          {sale.payment_method ??
                            '-'}
                        </span>
                      </div>

                      <span
                        className={
                          sale.status ===
                          'ANNULLATA'
                            ? 'badge badge-danger'
                            : 'badge badge-success'
                        }
                      >
                        {
                          sale.status
                        }
                      </span>
                    </div>

                    <div className="sale-items">
                      {sale.sale_items?.map(
                        (
                          item
                        ) => (
                          <div
                            key={
                              item.id
                            }
                          >
                            <span>
                              {getProductName(
                                item.products
                              )}
                              {' × '}
                              {
                                item.quantity
                              }
                            </span>

                            <strong>
                              {formatPrice(
                                item.line_total
                              )}
                            </strong>
                          </div>
                        )
                      )}
                    </div>

                    <div className="sale-total">
                      <span>
                        Totale
                      </span>

                      <strong>
                        {formatPrice(
                          sale.total_amount
                        )}
                      </strong>
                    </div>

                    {sale.status ===
                      'COMPLETATA' && (
                      <button
                        className="btn btn-danger"
                        onClick={() =>
                          cancelSale(
                            sale.id
                          )
                        }
                      >
                        Annulla vendita
                      </button>
                    )}

                    {sale.status ===
                      'ANNULLATA' && (
                      <div className="cancel-reason">
                        Motivo:{' '}
                        {
                          sale.cancellation_reason
                        }
                      </div>
                    )}
                  </div>
                )
              )}
            </div>
          </>
        )}

        {page ===
          'movements' && (
          <>
            <PageTitle
              title="Storico movimenti"
              subtitle="Tutte le variazioni registrate nel magazzino"
            />

            <div className="panel">
              <div className="table-wrapper">
                <table>
                  <thead>
                    <tr>
                      <th>
                        Data
                      </th>
                      <th>
                        Prodotto
                      </th>
                      <th>
                        Movimento
                      </th>
                      <th>
                        Quantità
                      </th>
                      <th>
                        Prima
                      </th>
                      <th>
                        Dopo
                      </th>
                      <th>
                        Note
                      </th>
                    </tr>
                  </thead>

                  <tbody>
                    {movements.map(
                      (
                        movement
                      ) => (
                        <tr
                          key={
                            movement.id
                          }
                        >
                          <td>
                            {formatDate(
                              movement.created_at
                            )}
                          </td>

                          <td>
                            <strong>
                              {getProductName(
                                movement.products
                              )}
                            </strong>
                          </td>

                          <td>
                            <MovementBadge
                              type={
                                movement.movement_type
                              }
                            />
                          </td>

                          <td>
                            {
                              movement.quantity
                            }
                          </td>

                          <td>
                            {
                              movement.stock_before
                            }
                          </td>

                          <td>
                            {
                              movement.stock_after
                            }
                          </td>

                          <td>
                            {movement.notes ??
                              '-'}
                          </td>
                        </tr>
                      )
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </>
        )}

        {page === 'inactive' && (
          <>
            <PageTitle
              title="Prodotti disattivati"
              subtitle="Prodotti non più disponibili nelle operazioni correnti"
            />

            {inactiveProducts.length ===
            0 ? (
              <div className="panel empty-state">
                Non ci sono
                prodotti
                disattivati.
              </div>
            ) : (
              <div className="inactive-products-grid">

                {inactiveProducts.map(
                  (product) => (
                    <div
                      className="inactive-product-card"
                      key={
                        product.product_id
                      }
                    >
                      <div className="inactive-product-header">
                        <h3>
                          {
                            product.name
                          }
                        </h3>

                        <span className="badge badge-neutral">
                          DISATTIVATO
                        </span>
                      </div>

                      <div className="inactive-product-data">

                        <div>
                          <span>
                            Fornitore
                          </span>

                          <strong>
                            {product.supplier ??
                              '-'}
                          </strong>
                        </div>

                        <div>
                          <span>
                            Categoria
                          </span>

                          <strong>
                            {product.category ??
                              '-'}
                          </strong>
                        </div>

                        <div>
                          <span>
                            Disponibilità
                            residua
                          </span>

                          <strong>
                            {formatQuantity(
                              product.availability,
                              product.unit_type
                            )}
                          </strong>
                        </div>

                        <div>
                          <span>
                            Prezzo SOMS
                          </span>

                          <strong>
                            {formatPrice(
                              product.sale_price
                            )}
                          </strong>
                        </div>

                      </div>

                      <button
                        className="btn btn-primary btn-full"
                        onClick={() =>
                          reactivateProduct(
                            product
                          )
                        }
                      >
                        Riattiva prodotto
                      </button>
                    </div>
                  )
                )}

              </div>
            )}
          </>
        )}

        {errorMessage && (
          <div className="alert alert-error no-print">
            {errorMessage}
          </div>
        )}

        {successMessage && (
          <div className="alert alert-success no-print">
            {successMessage}
          </div>
        )}

      </main>

      {selectedProduct && (
        <div
          className="modal-overlay no-print"
          onMouseDown={() =>
            setSelectedProduct(
              null
            )
          }
        >
          <div
            className="product-modal"
            onMouseDown={(e) =>
              e.stopPropagation()
            }
          >
            <div className="product-modal-header">
              <div>
                <span className="modal-eyebrow">
                  Gestione prodotto
                </span>

                <h2>
                  {
                    selectedProduct.name
                  }
                </h2>
              </div>

              <button
                type="button"
                className="modal-close"
                onClick={() =>
                  setSelectedProduct(
                    null
                  )
                }
                aria-label="Chiudi"
              >
                ×
              </button>
            </div>

            <div className="product-modal-grid">

              <ProductDetail
                label="Fornitore"
                value={
                  selectedProduct.supplier ??
                  '-'
                }
              />

              <ProductDetail
                label="Categoria"
                value={
                  selectedProduct.category ??
                  '-'
                }
              />

              <ProductDetail
                label="Disponibilità"
                value={
                  formatQuantity(
                    selectedProduct.availability,
                    selectedProduct.unit_type
                  )
                }
              />

              <ProductDetail
                label="Unità di misura"
                value={
                  formatUnit(
                    selectedProduct.unit_type
                  )
                }
              />

              <ProductDetail
                label="Prezzo normale"
                value={
                  formatPrice(
                    selectedProduct.normal_price
                  )
                }
              />

              <ProductDetail
                label="Prezzo confidenziale"
                value={
                  formatPrice(
                    selectedProduct.confidential_price
                  )
                }
              />

              <ProductDetail
                label="Prezzo SOMS"
                value={
                  formatPrice(
                    selectedProduct.sale_price
                  )
                }
              />

              <div className="product-detail">
                <span>
                  Giorni alla
                  scadenza
                </span>

                <ExpirationBadge
                  days={
                    getDaysToExpiration(
                      selectedProduct.expiration_date
                    )
                  }
                />
              </div>

            </div>

            <div className="modal-expiration">
              <label>
                Data di scadenza
              </label>

              <input
                className="form-control"
                type="date"
                value={
                  selectedProduct.expiration_date ??
                  ''
                }
                onChange={(e) =>
                  updateExpirationDate(
                    selectedProduct.product_id,
                    e.target.value
                  )
                }
              />
            </div>

            <div className="modal-danger-zone">
              <div>
                <strong>
                  Disattiva prodotto
                </strong>

                <p>
                  Il prodotto verrà
                  nascosto dalle
                  operazioni correnti,
                  ma vendite e movimenti
                  storici resteranno
                  conservati.
                </p>
              </div>

              <button
                className="btn btn-danger"
                onClick={() =>
                  deactivateProduct(
                    selectedProduct
                  )
                }
              >
                Disattiva prodotto
              </button>
            </div>
          </div>
        </div>
      )}

    </div>
  )
}

function DifferenceBadge({
  difference,
  unit,
}: {
  difference:
    number | null
  unit: string
}) {
  if (
    difference === null
  ) {
    return (
      <span className="badge badge-neutral">
        -
      </span>
    )
  }

  if (
    difference === 0
  ) {
    return (
      <span className="badge badge-neutral">
        Nessuna
      </span>
    )
  }

  const sign =
    difference > 0
      ? '+'
      : ''

  return (
    <span
      className={
        difference > 0
          ? 'badge badge-success'
          : 'badge badge-danger'
      }
    >
      {sign}
      {formatQuantity(
        difference,
        unit
      )}
    </span>
  )
}

function ProductDetail({
  label,
  value,
}: {
  label: string
  value: string
}) {
  return (
    <div className="product-detail">
      <span>
        {label}
      </span>

      <strong>
        {value}
      </strong>
    </div>
  )
}

function CashflowChart({
  points,
}: {
  points:
    CashflowPoint[]
}) {
  if (
    points.length === 0
  ) {
    return (
      <div className="empty-state">
        Nessun dato disponibile.
      </div>
    )
  }

  const width = 1000
  const height = 360
  const padding = 55

  const values =
    points.map(
      (point) =>
        point.value
    )

  const min =
    Math.min(
      0,
      ...values
    )

  const max =
    Math.max(
      0,
      ...values
    )

  const range =
    max - min === 0
      ? 1
      : max - min

  const x = (
    index: number
  ) => {
    if (
      points.length ===
      1
    ) {
      return width / 2
    }

    return (
      padding +
      (index /
        (points.length -
          1)) *
        (width -
          padding * 2)
    )
  }

  const y = (
    value: number
  ) =>
    height -
    padding -
    ((value - min) /
      range) *
      (height -
        padding * 2)

  const polyline =
    points
      .map(
        (
          point,
          index
        ) =>
          `${x(
            index
          )},${y(
            point.value
          )}`
      )
      .join(' ')

  const zeroY =
    y(0)

  return (
    <div className="chart-wrapper">
      <svg
        viewBox={`0 0 ${width} ${height}`}
        className="cashflow-chart"
      >
        <line
          x1={padding}
          y1={zeroY}
          x2={
            width -
            padding
          }
          y2={zeroY}
          className="chart-zero-line"
        />

        <polyline
          points={
            polyline
          }
          className="chart-line"
        />

        {points.map(
          (
            point,
            index
          ) => (
            <circle
              key={`${point.date}-${index}`}
              cx={x(
                index
              )}
              cy={y(
                point.value
              )}
              r="5"
              className="chart-dot"
            >
              <title>
                {formatDate(
                  point.date
                )}
                {' — '}
                {formatPrice(
                  point.value
                )}
              </title>
            </circle>
          )
        )}

        <text
          x={padding}
          y={25}
          className="chart-label"
        >
          {formatPrice(
            max
          )}
        </text>

        <text
          x={padding}
          y={
            height -
            12
          }
          className="chart-label"
        >
          {formatPrice(
            min
          )}
        </text>
      </svg>
    </div>
  )
}

function ExpirationBadge({
  days,
}: {
  days:
    number | null
}) {
  if (days === null) {
    return (
      <span className="badge badge-neutral">
        -
      </span>
    )
  }

  if (days < 0) {
    return (
      <span className="badge badge-danger">
        Scaduto da{' '}
        {Math.abs(
          days
        )}{' '}
        gg
      </span>
    )
  }

  if (days <= 30) {
    return (
      <span className="badge badge-danger">
        {days} gg
      </span>
    )
  }

  if (days <= 90) {
    return (
      <span className="badge badge-warning">
        {days} gg
      </span>
    )
  }

  return (
    <span className="badge badge-success">
      {days} gg
    </span>
  )
}

function getDaysToExpiration(
  expirationDate:
    string | null
) {
  if (
    !expirationDate
  ) {
    return null
  }

  const today =
    new Date()

  today.setHours(
    0,
    0,
    0,
    0
  )

  const expiration =
    new Date(
      `${expirationDate}T00:00:00`
    )

  const difference =
    expiration.getTime() -
    today.getTime()

  return Math.ceil(
    difference /
      (1000 *
        60 *
        60 *
        24)
  )
}

function NavButton({
  label,
  active,
  onClick,
}: {
  label: string
  active: boolean
  onClick: () => void
}) {
  return (
    <button
      className={`nav-button ${
        active
          ? 'nav-button-active'
          : ''
      }`}
      onClick={
        onClick
      }
    >
      {label}
    </button>
  )
}

function PageTitle({
  title,
  subtitle,
}: {
  title: string
  subtitle: string
}) {
  return (
    <div className="page-title">
      <h2>
        {title}
      </h2>

      <p>
        {subtitle}
      </p>
    </div>
  )
}

function SummaryCard({
  label,
  value,
}: {
  label: string
  value:
    string | number
}) {
  return (
    <div className="summary-card">
      <span>
        {label}
      </span>

      <strong>
        {value}
      </strong>
    </div>
  )
}

function FormField({
  label,
  children,
}: {
  label: string
  children:
    React.ReactNode
}) {
  return (
    <label className="form-field">
      <span>
        {label}
      </span>

      {children}
    </label>
  )
}

function MovementBadge({
  type,
}: {
  type: string
}) {
  const positive = [
    'CARICO',
    'INVENTARIO_INIZIALE',
    'RESO',
    'RETTIFICA_POSITIVA',
    'ANNULLAMENTO_VENDITA',
  ].includes(type)

  const negative = [
    'VENDITA',
    'RETTIFICA_NEGATIVA',
    'DANNEGGIATO',
    'SCADUTO',
    'SCARTO',
  ].includes(type)

  return (
    <span
      className={
        positive
          ? 'badge badge-success'
          : negative
          ? 'badge badge-danger'
          : 'badge badge-neutral'
      }
    >
      {type.replaceAll(
        '_',
        ' '
      )}
    </span>
  )
}

function getProductName(
  products:
    ProductRelation
) {
  if (!products) {
    return '-'
  }

  if (
    Array.isArray(
      products
    )
  ) {
    return (
      products[0]
        ?.name ?? '-'
    )
  }

  return products.name
}

function formatPrice(
  value:
    number | null
) {
  if (
    value === null ||
    value === undefined
  ) {
    return '-'
  }

  return `${Number(
    value
  ).toLocaleString(
    'it-IT',
    {
      minimumFractionDigits:
        2,
      maximumFractionDigits:
        2,
    }
  )} €`
}

function formatDate(
  value: string
) {
  return new Date(
    value
  ).toLocaleString(
    'it-IT'
  )
}

function formatDateOnly(
  value: string
) {
  return new Date(
    `${value}T00:00:00`
  ).toLocaleDateString(
    'it-IT'
  )
}

function formatUnit(
  unit: string
) {
  if (
    unit === 'KG'
  ) {
    return 'kg'
  }

  if (
    unit === 'LITRO'
  ) {
    return 'litri'
  }

  return 'pezzi'
}

function formatQuantity(
  quantity: number,
  unit: string
) {
  const formatted =
    Number(
      quantity
    ).toLocaleString(
      'it-IT',
      {
        maximumFractionDigits:
          3,
      }
    )

  return `${formatted} ${formatUnit(
    unit
  )}`
}

export default App