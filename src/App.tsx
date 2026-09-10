import { useEffect, useMemo, useState } from 'react'
import { utils, writeFileXLSX } from 'xlsx'
import { supabase } from './lib/supabase'
import {
  exportInventoryExcel,
  exportMovementsExcel,
  exportSalesExcel,
  makeDateTimeFileName,
  parseInventoryExcel,
} from './utils/excel'
import type {
  CartItem,
  CashflowPoint,
  ExcelImportRow,
  ImportMode,
  InventoryRow,
  Movement,
  OrderItem,
  Page,
  ProductRelation,
  Sale,
} from './types/warehouse'
import './App.css'

function App() {
  const [page, setPage] = useState<Page>('inventory')
  const [inventory, setInventory] = useState<InventoryRow[]>([])
  const [inactiveProducts, setInactiveProducts] = useState<InventoryRow[]>([])
  const [selectedProduct, setSelectedProduct] = useState<InventoryRow | null>(null)

  const [cart, setCart] = useState<CartItem[]>([])
  const [orderItems, setOrderItems] = useState<OrderItem[]>([])
  const [sales, setSales] = useState<Sale[]>([])
  const [movements, setMovements] = useState<Movement[]>([])

  const [excelRows, setExcelRows] = useState<ExcelImportRow[]>([])
  const [excelFileName, setExcelFileName] = useState('')
  const [importMode, setImportMode] = useState<ImportMode>('update')
  const [importingExcel, setImportingExcel] = useState(false)

  const [userEmail, setUserEmail] = useState<string | null>(null)
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')

  const [paymentMethod, setPaymentMethod] = useState('CONTANTI')

  const [loadProductId, setLoadProductId] = useState('')
  const [loadQuantity, setLoadQuantity] = useState('')
  const [loadNotes, setLoadNotes] = useState('')

  const [showNewProduct, setShowNewProduct] = useState(false)
  const [newName, setNewName] = useState('')
  const [newSupplier, setNewSupplier] = useState('')
  const [newCategory, setNewCategory] = useState('')
  const [newNormalPrice, setNewNormalPrice] = useState('')
  const [newConfidentialPrice, setNewConfidentialPrice] = useState('')
  const [newSalePrice, setNewSalePrice] = useState('')
  const [newUnitType, setNewUnitType] = useState('PEZZO')
  const [newExpirationDate, setNewExpirationDate] = useState('')

  const [errorMessage, setErrorMessage] = useState<string | null>(null)
  const [successMessage, setSuccessMessage] = useState<string | null>(null)

  const allProducts = useMemo(
    () => [...inventory, ...inactiveProducts],
    [inventory, inactiveProducts]
  )

  const productsToDeactivate = useMemo(() => {
    if (importMode !== 'replace' || excelRows.length === 0) return []

    const importedIds = new Set(
      excelRows
        .map((row) => row.product_id)
        .filter((id): id is string => Boolean(id))
    )

    return inventory.filter((product) => !importedIds.has(product.product_id))
  }, [importMode, excelRows, inventory])

  const changedRows = useMemo(
    () =>
      excelRows.filter(
        (row) => row.action === 'NUOVO' || row.changes.length > 0
      ),
    [excelRows]
  )

  function clearMessages() {
    setErrorMessage(null)
    setSuccessMessage(null)
  }

  function parseDecimal(value: string): number | null {
    const cleaned = value.trim().replace(',', '.')

    if (!cleaned) return null

    const number = Number(cleaned)

    return Number.isNaN(number) ? null : number
  }

  async function loadInventory() {
    const { data: products, error: productsError } = await supabase
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
      setErrorMessage(productsError.message)
      return
    }

    const { data: stocks, error: stocksError } = await supabase
      .from('current_stock')
      .select('product_id, availability')

    if (stocksError) {
      setErrorMessage(stocksError.message)
      return
    }

    const stockMap = new Map(
      (stocks ?? []).map((stock) => [
        stock.product_id,
        Number(stock.availability),
      ])
    )

    setInventory(
      (products ?? []).map((item: any) => ({
        product_id: item.id,
        name: item.name,
        supplier: getProductName(item.suppliers),
        category: getProductName(item.categories),
        normal_price: item.normal_price,
        confidential_price: item.confidential_price,
        sale_price: item.sale_price,
        availability: stockMap.get(item.id) ?? 0,
        unit_type: item.unit_type ?? 'PEZZO',
        expiration_date: item.expiration_date ?? null,
      }))
    )
  }

  async function loadInactiveProducts() {
    const { data: products, error: productsError } = await supabase
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
      setErrorMessage(productsError.message)
      return
    }

    const { data: stocks, error: stocksError } = await supabase
      .from('current_stock')
      .select('product_id, availability')

    if (stocksError) {
      setErrorMessage(stocksError.message)
      return
    }

    const stockMap = new Map(
      (stocks ?? []).map((stock) => [
        stock.product_id,
        Number(stock.availability),
      ])
    )

    setInactiveProducts(
      (products ?? []).map((item: any) => ({
        product_id: item.id,
        name: item.name,
        supplier: getProductName(item.suppliers),
        category: getProductName(item.categories),
        normal_price: item.normal_price,
        confidential_price: item.confidential_price,
        sale_price: item.sale_price,
        availability: stockMap.get(item.id) ?? 0,
        unit_type: item.unit_type ?? 'PEZZO',
        expiration_date: item.expiration_date ?? null,
      }))
    )
  }

  async function loadSales() {
    const { data, error } = await supabase
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
      .order('created_at', { ascending: false })

    if (error) {
      setErrorMessage(error.message)
      return
    }

    setSales((data ?? []) as Sale[])
  }

  async function loadMovements() {
    const { data, error } = await supabase
      .from('stock_movements')
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
      .order('created_at', { ascending: false })

    if (error) {
      setErrorMessage(error.message)
      return
    }

    setMovements(
      (data ?? []).map((item: any) => ({
        id: item.id,
        movement_type: item.movement_type,
        quantity: Number(item.quantity),
        stock_before: Number(item.stock_before),
        stock_after: Number(item.stock_after),
        unit_cost:
          item.unit_cost === null ? null : Number(item.unit_cost),
        notes: item.notes,
        created_at: item.created_at,
        products: item.products ?? null,
      }))
    )
  }

  async function refreshAll() {
    await Promise.all([
      loadInventory(),
      loadInactiveProducts(),
      loadSales(),
      loadMovements(),
    ])
  }

  async function login() {
    clearMessages()

    const { error } = await supabase.auth.signInWithPassword({
      email,
      password,
    })

    if (error) {
      setErrorMessage(error.message)
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

    const { error } = await supabase
      .from('products')
      .update({
        expiration_date: value || null,
      })
      .eq('id', productId)

    if (error) {
      setErrorMessage(error.message)
      return
    }

    await loadInventory()

    setSelectedProduct((current) =>
      current?.product_id === productId
        ? {
            ...current,
            expiration_date: value || null,
          }
        : current
    )

    setSuccessMessage('Data di scadenza aggiornata')
  }

  async function deactivateProduct(product: InventoryRow) {
    clearMessages()

    if (
      !window.confirm(
        `Vuoi disattivare "${product.name}"?\n\nStorico vendite e movimenti resteranno conservati.`
      )
    ) {
      return
    }

    const { error } = await supabase
      .from('products')
      .update({
        active: false,
      })
      .eq('id', product.product_id)

    if (error) {
      setErrorMessage(error.message)
      return
    }

    setSelectedProduct(null)

    await refreshAll()

    setSuccessMessage(
      `${product.name} è stato disattivato`
    )
  }

  async function reactivateProduct(product: InventoryRow) {
    clearMessages()

    const { error } = await supabase
      .from('products')
      .update({
        active: true,
      })
      .eq('id', product.product_id)

    if (error) {
      setErrorMessage(error.message)
      return
    }

    await refreshAll()

    setSuccessMessage(
      `${product.name} è stato riattivato`
    )
  }

  function addToCart(product: InventoryRow) {
    clearMessages()

    if (product.sale_price === null) {
      setErrorMessage(
        'Questo prodotto non ha un prezzo di vendita'
      )
      return
    }

    if (product.availability <= 0) {
      setErrorMessage('Prodotto non disponibile')
      return
    }

    setCart((current) => {
      const existing = current.find(
        (item) => item.product_id === product.product_id
      )

      if (existing) {
        if (
          existing.quantity + 1 >
          product.availability
        ) {
          setErrorMessage(
            'Quantità superiore alla disponibilità'
          )

          return current
        }

        return current.map((item) =>
          item.product_id === product.product_id
            ? {
                ...item,
                quantity: item.quantity + 1,
              }
            : item
        )
      }

      return [
        ...current,
        {
          product_id: product.product_id,
          name: product.name,
          quantity: 1,
          sale_price: Number(product.sale_price),
        },
      ]
    })
  }

  function changeQuantity(
    productId: string,
    quantity: number
  ) {
    const product = inventory.find(
      (item) => item.product_id === productId
    )

    if (!product) return

    if (quantity <= 0) {
      setCart((current) =>
        current.filter(
          (item) => item.product_id !== productId
        )
      )

      return
    }

    if (quantity > product.availability) {
      setErrorMessage(
        'Quantità superiore alla disponibilità'
      )
      return
    }

    setCart((current) =>
      current.map((item) =>
        item.product_id === productId
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

    if (cart.length === 0) {
      setErrorMessage('Aggiungi almeno un prodotto')
      return
    }

    const { error } = await supabase.rpc(
      'register_sale',
      {
        p_items: cart.map((item) => ({
          product_id: item.product_id,
          quantity: item.quantity,
        })),
        p_payment_method: paymentMethod,
        p_notes: null,
      }
    )

    if (error) {
      setErrorMessage(error.message)
      return
    }

    setCart([])

    await refreshAll()

    setSuccessMessage(
      'Vendita registrata correttamente'
    )
  }

  async function cancelSale(saleId: string) {
    clearMessages()

    const reason = window.prompt(
      'Inserisci il motivo dell’annullamento:'
    )

    if (!reason?.trim()) return

    if (
      !window.confirm(
        'Confermi l’annullamento? La merce verrà restituita al magazzino.'
      )
    ) {
      return
    }

    const { error } = await supabase.rpc(
      'cancel_sale',
      {
        p_sale_id: saleId,
        p_reason: reason.trim(),
      }
    )

    if (error) {
      setErrorMessage(error.message)
      return
    }

    await refreshAll()

    setSuccessMessage(
      'Vendita annullata correttamente'
    )
  }

  async function registerLoad() {
    clearMessages()

    const quantity = parseDecimal(loadQuantity)

    if (!loadProductId) {
      setErrorMessage('Seleziona un prodotto')
      return
    }

    if (
      quantity === null ||
      quantity <= 0
    ) {
      setErrorMessage(
        'Inserisci una quantità valida maggiore di zero'
      )
      return
    }

    const { error } = await supabase.rpc(
      'register_stock_load',
      {
        p_product_id: loadProductId,
        p_quantity: quantity,
        p_notes: loadNotes.trim() || null,
      }
    )

    if (error) {
      setErrorMessage(error.message)
      return
    }

    setLoadQuantity('')
    setLoadNotes('')

    await refreshAll()

    setSuccessMessage(
      'Carico registrato correttamente'
    )
  }

  async function createNewProduct() {
    clearMessages()

    if (
      !newName.trim() ||
      !newSupplier.trim() ||
      !newCategory.trim()
    ) {
      setErrorMessage(
        'Nome, fornitore e categoria sono obbligatori'
      )
      return
    }

    const supplierId =
      await getOrCreateNamedRecord(
        'suppliers',
        newSupplier.trim()
      )

    if (!supplierId) return

    const categoryId =
      await getOrCreateNamedRecord(
        'categories',
        newCategory.trim()
      )

    if (!categoryId) return

    const { data: product, error } =
      await supabase
        .from('products')
        .insert({
          name: newName.trim(),
          supplier_id: supplierId,
          category_id: categoryId,
          normal_price:
            parseDecimal(newNormalPrice),
          confidential_price:
            parseDecimal(newConfidentialPrice),
          sale_price:
            parseDecimal(newSalePrice),
          unit_type: newUnitType,
          expiration_date:
            newExpirationDate || null,
          active: true,
        })
        .select('id')
        .single()

    if (error) {
      setErrorMessage(error.message)
      return
    }

    setLoadProductId(product.id)

    setNewName('')
    setNewSupplier('')
    setNewCategory('')
    setNewNormalPrice('')
    setNewConfidentialPrice('')
    setNewSalePrice('')
    setNewUnitType('PEZZO')
    setNewExpirationDate('')
    setShowNewProduct(false)

    await loadInventory()

    setSuccessMessage(
      'Prodotto creato. Ora inserisci la quantità ricevuta.'
    )
  }

  async function getOrCreateNamedRecord(
    table: 'suppliers' | 'categories',
    name: string
  ): Promise<string | null> {
    const {
      data: existing,
      error: searchError,
    } = await supabase
      .from(table)
      .select('id')
      .eq('name', name)
      .maybeSingle()

    if (searchError) {
      setErrorMessage(searchError.message)
      return null
    }

    if (existing) {
      return existing.id
    }

    const {
      data: created,
      error,
    } = await supabase
      .from(table)
      .insert({
        name,
      })
      .select('id')
      .single()

    if (error) {
      setErrorMessage(error.message)
      return null
    }

    return created.id
  }

  function changeOrderQuantity(
    productId: string,
    value: string
  ) {
    setOrderItems((current) => {
      const existing = current.find(
        (item) => item.product_id === productId
      )

      if (existing) {
        return current.map((item) =>
          item.product_id === productId
            ? {
                ...item,
                quantity: value,
              }
            : item
        )
      }

      return [
        ...current,
        {
          product_id: productId,
          quantity: value,
        },
      ]
    })
  }

  function getOrderQuantity(productId: string) {
    return (
      orderItems.find(
        (item) => item.product_id === productId
      )?.quantity ?? ''
    )
  }

  function getPurchasePrice(
    product: InventoryRow
  ) {
    const confidential = Number(
      product.confidential_price ?? 0
    )

    return confidential > 0
      ? confidential
      : Number(product.normal_price ?? 0)
  }

  function exportOrderExcel() {
    clearMessages()

    const rows = orderItems
      .map((orderItem) => {
        const product = inventory.find(
          (item) =>
            item.product_id ===
            orderItem.product_id
        )

        const quantity = parseDecimal(
          orderItem.quantity
        )

        if (
          !product ||
          quantity === null ||
          quantity <= 0
        ) {
          return null
        }

        const unitPrice =
          getPurchasePrice(product)

        return {
          Prodotto: product.name,
          Fornitore:
            product.supplier ?? '-',
          Quantità: quantity,
          Unità: formatUnit(
            product.unit_type
          ),
          'Prezzo unitario (€)':
            unitPrice,
          'Spesa prevista (€)': Number(
            (
              quantity *
              unitPrice
            ).toFixed(2)
          ),
        }
      })
      .filter(Boolean) as Record<
        string,
        string | number
      >[]

    if (rows.length === 0) {
      setErrorMessage(
        'Inserisci almeno una quantità da ordinare'
      )
      return
    }

    const total = rows.reduce(
      (sum, row) =>
        sum +
        Number(
          row['Spesa prevista (€)']
        ),
      0
    )

    const worksheet =
      utils.json_to_sheet([
        ...rows,
        {},
        {
          Prodotto: 'TOTALE',
          'Spesa prevista (€)': total,
        },
      ])

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
      makeDateTimeFileName('Ordini')

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

  async function handleExcelUpload(
    file: File,
    mode: ImportMode
  ) {
    clearMessages()

    try {
      const rows =
        await parseInventoryExcel(
          file,
          allProducts
        )

      setExcelRows(rows)
      setExcelFileName(file.name)
      setImportMode(mode)

      setSuccessMessage(
        'Excel letto correttamente. Controlla il riquadro delle modifiche prima di confermare.'
      )
    } catch (error) {
      setExcelRows([])
      setExcelFileName('')

      setErrorMessage(
        error instanceof Error
          ? error.message
          : 'Errore durante la lettura del file Excel'
      )
    }
  }

  async function confirmExcelImport() {
    clearMessages()

    if (excelRows.length === 0) {
      setErrorMessage(
        'Non ci sono righe da importare'
      )
      return
    }

    if (importMode === 'replace') {
      const firstConfirm =
        window.confirm(
          `ATTENZIONE\n\nStai per sostituire l'inventario corrente con "${excelFileName}".\n\n` +
            `${productsToDeactivate.length} prodotti attualmente attivi e assenti dal file verranno azzerati e disattivati.\n` +
            `Lo storico vendite e movimenti verrà conservato.\n\nVuoi continuare?`
        )

      if (!firstConfirm) return

      const typed = window.prompt(
        'Per confermare la sostituzione scrivi esattamente: SOSTITUISCI'
      )

      if (typed !== 'SOSTITUISCI') {
        setErrorMessage(
          'Sostituzione annullata: conferma non valida'
        )
        return
      }
    } else {
      if (
        !window.confirm(
          `Confermi l'aggiornamento di ${excelRows.length} righe?`
        )
      ) {
        return
      }
    }

    setImportingExcel(true)

    const payload = excelRows.map(
      (row) => ({
        product_id: row.product_id,
        name: row.name,
        supplier: row.supplier,
        category: row.category,
        normal_price:
          row.normal_price,
        confidential_price:
          row.confidential_price,
        sale_price: row.sale_price,
        availability:
          row.availability,
        unit_type: row.unit_type,
        expiration_date:
          row.expiration_date,
      })
    )

    const {
      data,
      error,
    } = await supabase.rpc(
      'import_inventory_database',
      {
        p_rows: payload,
        p_replace_all:
          importMode === 'replace',
      }
    )

    setImportingExcel(false)

    if (error) {
      setErrorMessage(error.message)
      return
    }

    setExcelRows([])
    setExcelFileName('')

    await refreshAll()

    const result = data as {
      created?: number
      updated?: number
      adjustments?: number
      deactivated?: number
    } | null

    setSuccessMessage(
      `Importazione completata: ${
        result?.created ?? 0
      } nuovi, ${
        result?.updated ?? 0
      } aggiornati, ${
        result?.adjustments ?? 0
      } rettifiche, ${
        result?.deactivated ?? 0
      } disattivati.`
    )
  }

  const selectedLoadProduct =
    inventory.find(
      (product) =>
        product.product_id ===
        loadProductId
    )

  const cartTotal = useMemo(
    () =>
      cart.reduce(
        (sum, item) =>
          sum +
          item.quantity *
            item.sale_price,
        0
      ),
    [cart]
  )

  const categoryCount = useMemo(
    () =>
      new Set(
        inventory
          .map((item) => item.category)
          .filter(Boolean)
      ).size,
    [inventory]
  )

  const orderTotal = useMemo(
    () =>
      orderItems.reduce(
        (sum, orderItem) => {
          const product =
            inventory.find(
              (item) =>
                item.product_id ===
                orderItem.product_id
            )

          const quantity =
            parseDecimal(
              orderItem.quantity
            )

          if (
            !product ||
            quantity === null ||
            quantity <= 0
          ) {
            return sum
          }

          return (
            sum +
            quantity *
              getPurchasePrice(product)
          )
        },
        0
      ),
    [orderItems, inventory]
  )

  const tiedCapital = useMemo(
    () =>
      allProducts.reduce(
        (sum, product) =>
          sum +
          product.availability *
            getPurchasePrice(product),
        0
      ),
    [allProducts]
  )

  const inventorySaleValue =
    useMemo(
      () =>
        allProducts.reduce(
          (sum, product) =>
            sum +
            product.availability *
              Number(
                product.sale_price ?? 0
              ),
          0
        ),
      [allProducts]
    )

  const completedSalesRevenue =
    useMemo(
      () =>
        sales
          .filter(
            (sale) =>
              sale.status ===
              'COMPLETATA'
          )
          .reduce(
            (sum, sale) =>
              sum +
              Number(
                sale.total_amount
              ),
            0
          ),
      [sales]
    )

  const purchaseOutflow =
    useMemo(
      () =>
        movements
          .filter((movement) =>
            [
              'INVENTARIO_INIZIALE',
              'CARICO',
            ].includes(
              movement.movement_type
            )
          )
          .reduce(
            (sum, movement) =>
              sum +
              movement.quantity *
                Number(
                  movement.unit_cost ??
                    0
                ),
            0
          ),
      [movements]
    )

  const cashflowPoints =
    useMemo(() => {
      const events: {
        date: string
        amount: number
      }[] = []

      movements
        .filter((movement) =>
          [
            'INVENTARIO_INIZIALE',
            'CARICO',
          ].includes(
            movement.movement_type
          )
        )
        .forEach((movement) =>
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
        )

      sales
        .filter(
          (sale) =>
            sale.status ===
            'COMPLETATA'
        )
        .forEach((sale) =>
          events.push({
            date: sale.created_at,
            amount: Number(
              sale.total_amount
            ),
          })
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
          cumulative += event.amount

          return {
            date: event.date,
            value: cumulative,
          }
        }
      )
    }, [movements, sales])

  useEffect(() => {
    supabase.auth
      .getSession()
      .then(({ data }) => {
        setUserEmail(
          data.session?.user.email ??
            null
        )
      })

    const { data } =
      supabase.auth.onAuthStateChange(
        (_event, session) => {
          setUserEmail(
            session?.user.email ??
              null
          )
        }
      )

    return () =>
      data.subscription.unsubscribe()
  }, [])

  useEffect(() => {
    if (userEmail) {
      refreshAll()
    } else {
      setInventory([])
      setInactiveProducts([])
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
            Gramsci Warehouse Management
          </h1>

          <p>
            Accedi per gestire il magazzino.
          </p>

          <input
            className="form-control"
            type="email"
            placeholder="Email"
            value={email}
            onChange={(e) =>
              setEmail(e.target.value)
            }
          />

          <input
            className="form-control"
            type="password"
            placeholder="Password"
            value={password}
            onChange={(e) =>
              setPassword(e.target.value)
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
        <img
          src="/logo.png"
          alt="Logo"
          className="logo-image"
        />

        <div className="header-title">
          <h1>
            Gramsci Warehouse Management
          </h1>

          <span>
            Gestione magazzino
          </span>
        </div>

        <div className="user-area">
          <span>{userEmail}</span>

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
          pageName="inventory"
          page={page}
          setPage={setPage}
        />

        <NavButton
          label="Vendite"
          pageName="sales"
          page={page}
          setPage={setPage}
        />

        <NavButton
          label="Carico merce"
          pageName="load"
          page={page}
          setPage={setPage}
        />

        <NavButton
          label="Ordina"
          pageName="order"
          page={page}
          setPage={setPage}
        />

        <NavButton
          label="Statistiche"
          pageName="statistics"
          page={page}
          setPage={setPage}
        />

        <NavButton
          label="Storico vendite"
          pageName="salesHistory"
          page={page}
          setPage={setPage}
        />

        <NavButton
          label="Movimenti"
          pageName="movements"
          page={page}
          setPage={setPage}
        />

        <NavButton
          label="Prodotti disattivati"
          pageName="inactive"
          page={page}
          setPage={setPage}
        />
      </nav>

      <main className="main-content">
        {page === 'inventory' && (
          <section className="inventory-print-area">
            <div className="print-only print-inventory-header">
              <h1>
                Gramsci Warehouse Management
              </h1>

              <h2>Inventario</h2>

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
                  onClick={() => {
                    clearMessages()

                    if (!inventory.length) {
                      setErrorMessage(
                        'L’inventario è vuoto'
                      )
                      return
                    }

                    const filename =
                      exportInventoryExcel(
                        inventory
                      )

                    setSuccessMessage(
                      `File ${filename} generato correttamente`
                    )
                  }}
                >
                  Esporta Excel
                </button>

                <button
                  className="btn btn-secondary"
                  onClick={() =>
                    window.print()
                  }
                >
                  Stampa inventario
                </button>

                <label className="btn btn-upload">
                  Aggiorna da Excel

                  <input
                    className="hidden-file-input"
                    type="file"
                    accept=".xlsx,.xls"
                    onChange={async (
                      e
                    ) => {
                      const file =
                        e.target.files?.[0]

                      if (file) {
                        await handleExcelUpload(
                          file,
                          'update'
                        )
                      }

                      e.currentTarget.value =
                        ''
                    }}
                  />
                </label>

                <label className="btn btn-danger-solid">
                  Importa database

                  <input
                    className="hidden-file-input"
                    type="file"
                    accept=".xlsx,.xls"
                    onChange={async (
                      e
                    ) => {
                      const file =
                        e.target.files?.[0]

                      if (file) {
                        await handleExcelUpload(
                          file,
                          'replace'
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
                  value={inventory.length}
                />

                <SummaryCard
                  label="Tipologie"
                  value={categoryCount}
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

              {excelRows.length > 0 && (
                <ImportPreview
                  mode={importMode}
                  fileName={excelFileName}
                  rows={excelRows}
                  changedRows={
                    changedRows
                  }
                  productsToDeactivate={
                    productsToDeactivate
                  }
                  importing={
                    importingExcel
                  }
                  onCancel={() => {
                    setExcelRows([])
                    setExcelFileName('')
                    clearMessages()
                  }}
                  onConfirm={
                    confirmExcelImport
                  }
                />
              )}
            </div>

            <InventoryTable
              inventory={inventory}
              onSelect={
                setSelectedProduct
              }
              onExpirationChange={
                updateExpirationDate
              }
            />
          </section>
        )}

        {page === 'sales' && (
          <>
            <PageTitle
              title="Nuova vendita"
              subtitle="Seleziona i prodotti e registra la vendita"
            />

            <div className="two-column-layout">
              <div className="panel sales-products-panel">
                <h3 className="sales-products-title">
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
                            {product.name}
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
                            addToCart(product)
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
                <h3>Carrello</h3>

                {!cart.length && (
                  <div className="empty-state">
                    Nessun prodotto nel
                    carrello.
                  </div>
                )}

                {cart.map((item) => (
                  <div
                    className="cart-row"
                    key={item.product_id}
                  >
                    <div>
                      <strong>
                        {item.name}
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
                        onChange={(e) =>
                          changeQuantity(
                            item.product_id,
                            Number(
                              e.target.value
                            )
                          )
                        }
                      />

                      <button
                        className="btn btn-danger btn-small"
                        onClick={() =>
                          setCart(
                            (current) =>
                              current.filter(
                                (row) =>
                                  row.product_id !==
                                  item.product_id
                              )
                          )
                        }
                      >
                        Rimuovi
                      </button>
                    </div>
                  </div>
                ))}

                <div className="checkout-total">
                  <span>Totale</span>

                  <strong>
                    {formatPrice(
                      cartTotal
                    )}
                  </strong>
                </div>

                <label>
                  Metodo di pagamento
                </label>

                <select
                  className="form-control"
                  value={paymentMethod}
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
                onClick={() =>
                  setShowNewProduct(
                    (value) => !value
                  )
                }
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
                        value={newName}
                        onChange={(e) =>
                          setNewName(
                            e.target.value
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
                        onChange={(e) =>
                          setNewSupplier(
                            e.target.value
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
                        onChange={(e) =>
                          setNewCategory(
                            e.target.value
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
                        onChange={(e) =>
                          setNewUnitType(
                            e.target.value
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
                        onChange={(e) =>
                          setNewNormalPrice(
                            e.target.value
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
                        onChange={(e) =>
                          setNewConfidentialPrice(
                            e.target.value
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
                        onChange={(e) =>
                          setNewSalePrice(
                            e.target.value
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
                        onChange={(e) =>
                          setNewExpirationDate(
                            e.target.value
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

                    setLoadQuantity('')
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
                        {product.name}
                      </option>
                    )
                  )}
                </select>
              </FormField>

              {selectedLoadProduct && (
                <div className="product-info-card">
                  <Info
                    label="Fornitore"
                    value={
                      selectedLoadProduct.supplier ??
                      '-'
                    }
                  />

                  <Info
                    label="Categoria"
                    value={
                      selectedLoadProduct.category ??
                      '-'
                    }
                  />

                  <Info
                    label="Disponibilità"
                    value={formatQuantity(
                      selectedLoadProduct.availability,
                      selectedLoadProduct.unit_type
                    )}
                  />
                </div>
              )}

              <FormField label="Quantità ricevuta">
                <input
                  className="form-control"
                  type="text"
                  inputMode="decimal"
                  value={loadQuantity}
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
                  value={loadNotes}
                  onChange={(e) =>
                    setLoadNotes(
                      e.target.value
                    )
                  }
                />
              </FormField>

              <button
                className="btn btn-primary"
                onClick={registerLoad}
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
                      <th>Prodotto</th>
                      <th>Fornitore</th>
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
                          quantity !== null &&
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
                                {product.name}
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
                                value={
                                  quantityText
                                }
                                onChange={(e) =>
                                  changeOrderQuantity(
                                    product.product_id,
                                    e.target.value
                                  )
                                }
                              />
                            </td>

                            <td>
                              {rowTotal > 0
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
                    Spesa totale prevista
                  </span>

                  <strong>
                    {formatPrice(
                      orderTotal
                    )}
                  </strong>
                </div>

                <button
                  className="btn btn-primary"
                  onClick={exportOrderExcel}
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
                value={formatPrice(
                  tiedCapital
                )}
              />

              <SummaryCard
                label="Valore potenziale di vendita"
                value={formatPrice(
                  inventorySaleValue
                )}
              />

              <SummaryCard
                label="Margine potenziale"
                value={formatPrice(
                  inventorySaleValue -
                    tiedCapital
                )}
              />

              <SummaryCard
                label="Ricavi vendite"
                value={formatPrice(
                  completedSalesRevenue
                )}
              />

              <SummaryCard
                label="Capitale investito storico"
                value={formatPrice(
                  purchaseOutflow
                )}
              />

              <SummaryCard
                label="Cashflow cumulato"
                value={formatPrice(
                  completedSalesRevenue -
                    purchaseOutflow
                )}
              />
            </div>

            <div className="panel">
              <h3>
                Cashflow cumulato
              </h3>

              <CashflowChart
                points={cashflowPoints}
              />
            </div>
          </>
        )}

        {page === 'salesHistory' && (
          <>
            <PageTitle
              title="Storico vendite"
              subtitle="Vendite registrate e annullamenti"
            />

            <div className="page-actions">
              <button
                className="btn btn-primary"
                onClick={() => {
                  if (!sales.length) {
                    setErrorMessage(
                      'Non ci sono vendite da esportare'
                    )
                    return
                  }

                  const filename =
                    exportSalesExcel(
                      sales
                    )

                  setSuccessMessage(
                    `File ${filename} generato correttamente`
                  )
                }}
              >
                Esporta storico vendite
              </button>
            </div>

            <div className="cards-list">
              {sales.map((sale) => (
                <div
                  key={sale.id}
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
                      {sale.status}
                    </span>
                  </div>

                  <div className="sale-items">
                    {sale.sale_items?.map(
                      (item) => (
                        <div
                          key={item.id}
                        >
                          <span>
                            {getProductName(
                              item.products
                            )}
                            {' × '}
                            {item.quantity}
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
                    <span>Totale</span>

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
              ))}
            </div>
          </>
        )}

        {page === 'movements' && (
          <>
            <PageTitle
              title="Storico movimenti"
              subtitle="Tutte le variazioni registrate nel magazzino"
            />

            <div className="page-actions">
              <button
                className="btn btn-primary"
                onClick={() => {
                  if (
                    !movements.length
                  ) {
                    setErrorMessage(
                      'Non ci sono movimenti da esportare'
                    )
                    return
                  }

                  const filename =
                    exportMovementsExcel(
                      movements
                    )

                  setSuccessMessage(
                    `File ${filename} generato correttamente`
                  )
                }}
              >
                Esporta movimenti
              </button>
            </div>

            <div className="panel">
              <div className="table-wrapper">
                <table>
                  <thead>
                    <tr>
                      <th>Data</th>
                      <th>Prodotto</th>
                      <th>Movimento</th>
                      <th>Quantità</th>
                      <th>Prima</th>
                      <th>Dopo</th>
                      <th>
                        Costo unitario
                      </th>
                      <th>Note</th>
                    </tr>
                  </thead>

                  <tbody>
                    {movements.map(
                      (movement) => (
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
                            {formatPrice(
                              movement.unit_cost
                            )}
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

            {!inactiveProducts.length ? (
              <div className="panel empty-state">
                Non ci sono prodotti
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
                          {product.name}
                        </h3>

                        <span className="badge badge-neutral">
                          DISATTIVATO
                        </span>
                      </div>

                      <div className="inactive-product-data">
                        <Info
                          label="Fornitore"
                          value={
                            product.supplier ??
                            '-'
                          }
                        />

                        <Info
                          label="Categoria"
                          value={
                            product.category ??
                            '-'
                          }
                        />

                        <Info
                          label="Disponibilità residua"
                          value={formatQuantity(
                            product.availability,
                            product.unit_type
                          )}
                        />

                        <Info
                          label="Prezzo SOMS"
                          value={formatPrice(
                            product.sale_price
                          )}
                        />
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
        <ProductModal
          product={selectedProduct}
          onClose={() =>
            setSelectedProduct(null)
          }
          onExpirationChange={
            updateExpirationDate
          }
          onDeactivate={
            deactivateProduct
          }
        />
      )}
    </div>
  )
}

function ImportPreview({
  mode,
  fileName,
  rows,
  changedRows,
  productsToDeactivate,
  importing,
  onCancel,
  onConfirm,
}: {
  mode: ImportMode
  fileName: string
  rows: ExcelImportRow[]
  changedRows: ExcelImportRow[]
  productsToDeactivate: InventoryRow[]
  importing: boolean
  onCancel: () => void
  onConfirm: () => void
}) {
  const newCount = rows.filter(
    (row) => row.action === 'NUOVO'
  ).length

  const adjustmentCount =
    rows.filter(
      (row) =>
        row.difference !== null &&
        row.difference !== 0
    ).length

  return (
    <div
      className={`excel-import-panel ${
        mode === 'replace'
          ? 'excel-import-danger'
          : ''
      }`}
    >
      <div className="excel-import-header">
        <div>
          <span className="modal-eyebrow">
            {mode === 'replace'
              ? 'Sostituzione inventario'
              : 'Aggiornamento inventario'}
          </span>

          <h3>{fileName}</h3>
        </div>

        <span
          className={
            mode === 'replace'
              ? 'badge badge-danger'
              : 'badge badge-warning'
          }
        >
          Nessuna modifica ancora
          applicata
        </span>
      </div>

      <div className="import-summary-grid">
        <SummaryCard
          label="Righe Excel"
          value={rows.length}
        />

        <SummaryCard
          label="Nuovi prodotti"
          value={newCount}
        />

        <SummaryCard
          label="Righe con modifiche"
          value={changedRows.length}
        />

        <SummaryCard
          label="Rettifiche giacenza"
          value={adjustmentCount}
        />

        {mode === 'replace' && (
          <SummaryCard
            label="Da disattivare"
            value={
              productsToDeactivate.length
            }
          />
        )}
      </div>

      {changedRows.length === 0 &&
        (mode !== 'replace' ||
          productsToDeactivate.length ===
            0) && (
          <div className="no-changes-box">
            Il file coincide con
            l’inventario attuale: non
            risultano modifiche.
          </div>
        )}

      {changedRows.length > 0 && (
        <>
          <h4>Cosa cambierà</h4>

          <div className="changes-list">
            {changedRows.map(
              (row, index) => (
                <div
                  className="change-card"
                  key={`${
                    row.product_id ??
                    row.name
                  }-${index}`}
                >
                  <div className="change-card-title">
                    <strong>
                      {row.name}
                    </strong>

                    <span
                      className={
                        row.action ===
                        'NUOVO'
                          ? 'badge badge-warning'
                          : 'badge badge-neutral'
                      }
                    >
                      {row.action}
                    </span>
                  </div>

                  <ul>
                    {row.changes.map(
                      (change) => (
                        <li key={change}>
                          {change}
                        </li>
                      )
                    )}
                  </ul>
                </div>
              )
            )}
          </div>
        </>
      )}

      {mode === 'replace' &&
        productsToDeactivate.length >
          0 && (
          <div className="deactivate-preview">
            <h4>
              Prodotti assenti
              dall’Excel
            </h4>

            <p>
              Questi prodotti verranno
              portati a giacenza zero
              tramite rettifica e poi
              disattivati:
            </p>

            <div className="deactivate-tags">
              {productsToDeactivate.map(
                (product) => (
                  <span
                    key={
                      product.product_id
                    }
                    className="badge badge-danger"
                  >
                    {product.name}
                  </span>
                )
              )}
            </div>
          </div>
        )}

      <div className="import-actions">
        <button
          className="btn btn-secondary"
          onClick={onCancel}
          disabled={importing}
        >
          Annulla
        </button>

        <button
          className={
            mode === 'replace'
              ? 'btn btn-danger-solid btn-large'
              : 'btn btn-primary btn-large'
          }
          onClick={onConfirm}
          disabled={importing}
        >
          {importing
            ? 'Importazione...'
            : mode === 'replace'
            ? 'Sostituisci inventario'
            : 'Conferma aggiornamento'}
        </button>
      </div>
    </div>
  )
}

function InventoryTable({
  inventory,
  onSelect,
  onExpirationChange,
}: {
  inventory: InventoryRow[]
  onSelect: (
    product: InventoryRow
  ) => void
  onExpirationChange: (
    productId: string,
    value: string
  ) => void
}) {
  return (
    <div className="panel inventory-panel">
      <div className="table-wrapper">
        <table className="inventory-table">
          <thead>
            <tr>
              <th>Prodotto</th>
              <th>Fornitore</th>
              <th>Tipo</th>
              <th>Prezzo</th>
              <th>
                P. confidenziale
              </th>
              <th>P. SOMS</th>
              <th>
                Disponibilità
              </th>
              <th>Scadenza</th>
              <th>
                Giorni residui
              </th>
            </tr>
          </thead>

          <tbody>
            {inventory.map((item) => (
              <tr
                key={item.product_id}
              >
                <td>
                  <button
                    type="button"
                    className="product-name-button no-print"
                    onClick={() =>
                      onSelect(item)
                    }
                  >
                    {item.name}
                  </button>

                  <span className="print-only">
                    {item.name}
                  </span>
                </td>

                <td>
                  {item.supplier ?? '-'}
                </td>

                <td>
                  {item.category ?? '-'}
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
                    onChange={(e) =>
                      onExpirationChange(
                        item.product_id,
                        e.target.value
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
                    days={getDaysToExpiration(
                      item.expiration_date
                    )}
                  />
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  )
}

function ProductModal({
  product,
  onClose,
  onExpirationChange,
  onDeactivate,
}: {
  product: InventoryRow
  onClose: () => void
  onExpirationChange: (
    productId: string,
    value: string
  ) => void
  onDeactivate: (
    product: InventoryRow
  ) => void
}) {
  return (
    <div
      className="modal-overlay no-print"
      onMouseDown={onClose}
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

            <h2>{product.name}</h2>
          </div>

          <button
            type="button"
            className="modal-close"
            onClick={onClose}
            aria-label="Chiudi"
          >
            ×
          </button>
        </div>

        <div className="product-modal-grid">
          <ProductDetail
            label="Fornitore"
            value={
              product.supplier ?? '-'
            }
          />

          <ProductDetail
            label="Categoria"
            value={
              product.category ?? '-'
            }
          />

          <ProductDetail
            label="Disponibilità"
            value={formatQuantity(
              product.availability,
              product.unit_type
            )}
          />

          <ProductDetail
            label="Unità"
            value={formatUnit(
              product.unit_type
            )}
          />

          <ProductDetail
            label="Prezzo normale"
            value={formatPrice(
              product.normal_price
            )}
          />

          <ProductDetail
            label="Prezzo confidenziale"
            value={formatPrice(
              product.confidential_price
            )}
          />

          <ProductDetail
            label="Prezzo SOMS"
            value={formatPrice(
              product.sale_price
            )}
          />

          <div className="product-detail">
            <span>
              Giorni alla scadenza
            </span>

            <ExpirationBadge
              days={getDaysToExpiration(
                product.expiration_date
              )}
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
              product.expiration_date ??
              ''
            }
            onChange={(e) =>
              onExpirationChange(
                product.product_id,
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
              Il prodotto verrà nascosto
              dalle operazioni correnti,
              ma vendite e movimenti
              storici resteranno
              conservati.
            </p>
          </div>

          <button
            className="btn btn-danger"
            onClick={() =>
              onDeactivate(product)
            }
          >
            Disattiva prodotto
          </button>
        </div>
      </div>
    </div>
  )
}

function CashflowChart({
  points,
}: {
  points: CashflowPoint[]
}) {
  if (!points.length) {
    return (
      <div className="empty-state">
        Nessun dato disponibile.
      </div>
    )
  }

  const width = 1000
  const height = 360
  const padding = 55

  const values = points.map(
    (point) => point.value
  )

  const min = Math.min(
    0,
    ...values
  )

  const max = Math.max(
    0,
    ...values
  )

  const range =
    max - min === 0
      ? 1
      : max - min

  const x = (index: number) =>
    points.length === 1
      ? width / 2
      : padding +
        (index /
          (points.length - 1)) *
          (width -
            padding * 2)

  const y = (value: number) =>
    height -
    padding -
    ((value - min) / range) *
      (height -
        padding * 2)

  const polyline = points
    .map(
      (point, index) =>
        `${x(index)},${y(
          point.value
        )}`
    )
    .join(' ')

  return (
    <div className="chart-wrapper">
      <svg
        viewBox={`0 0 ${width} ${height}`}
        className="cashflow-chart"
      >
        <line
          x1={padding}
          y1={y(0)}
          x2={width - padding}
          y2={y(0)}
          className="chart-zero-line"
        />

        <polyline
          points={polyline}
          className="chart-line"
        />

        {points.map(
          (point, index) => (
            <circle
              key={`${point.date}-${index}`}
              cx={x(index)}
              cy={y(point.value)}
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
      </svg>
    </div>
  )
}

function NavButton({
  label,
  pageName,
  page,
  setPage,
}: {
  label: string
  pageName: Page
  page: Page
  setPage: (page: Page) => void
}) {
  return (
    <button
      className={`nav-button ${
        page === pageName
          ? 'nav-button-active'
          : ''
      }`}
      onClick={() =>
        setPage(pageName)
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
      <h2>{title}</h2>
      <p>{subtitle}</p>
    </div>
  )
}

function SummaryCard({
  label,
  value,
}: {
  label: string
  value: string | number
}) {
  return (
    <div className="summary-card">
      <span>{label}</span>
      <strong>{value}</strong>
    </div>
  )
}

function FormField({
  label,
  children,
}: {
  label: string
  children: React.ReactNode
}) {
  return (
    <label className="form-field">
      <span>{label}</span>
      {children}
    </label>
  )
}

function Info({
  label,
  value,
}: {
  label: string
  value: string
}) {
  return (
    <div>
      <span>{label}</span>
      <strong>{value}</strong>
    </div>
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
      <span>{label}</span>
      <strong>{value}</strong>
    </div>
  )
}

function ExpirationBadge({
  days,
}: {
  days: number | null
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
        Scaduto da {Math.abs(days)} gg
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
      {type.replaceAll('_', ' ')}
    </span>
  )
}

function getDaysToExpiration(
  expirationDate: string | null
) {
  if (!expirationDate) return null

  const today = new Date()

  today.setHours(
    0,
    0,
    0,
    0
  )

  const expiration = new Date(
    `${expirationDate}T00:00:00`
  )

  return Math.ceil(
    (
      expiration.getTime() -
      today.getTime()
    ) /
      (1000 *
        60 *
        60 *
        24)
  )
}

function getProductName(
  products: ProductRelation
) {
  if (!products) return '-'

  if (Array.isArray(products)) {
    return products[0]?.name ?? '-'
  }

  return products.name
}

function formatPrice(
  value: number | null
) {
  if (
    value === null ||
    value === undefined
  ) {
    return '-'
  }

  return `${Number(
    value
  ).toLocaleString('it-IT', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  })} €`
}

function formatDate(value: string) {
  return new Date(
    value
  ).toLocaleString('it-IT')
}

function formatDateOnly(
  value: string
) {
  return new Date(
    `${value}T00:00:00`
  ).toLocaleDateString('it-IT')
}

function formatUnit(unit: string) {
  if (unit === 'KG') return 'kg'

  if (unit === 'LITRO') {
    return 'litri'
  }

  return 'pezzi'
}

function formatQuantity(
  quantity: number,
  unit: string
) {
  return `${Number(
    quantity
  ).toLocaleString('it-IT', {
    maximumFractionDigits: 3,
  })} ${formatUnit(unit)}`
}

export default App