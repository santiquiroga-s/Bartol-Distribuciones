import React, { useEffect, useMemo, useState } from 'react';
import ExcelJS from 'exceljs';
import { createRoot } from 'react-dom/client';
import { supabase } from './lib/supabase';
import bartolLogo from './assets/bartol-logo.png';
import huevoSkate from './assets/huevo-skate.png';
import {
  Search,
  Plus,
  Pencil,
  Trash2,
  X,
  Package,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  Sparkles,
  LayoutGrid,
  List,
  Download,
  CheckCircle2,
  AlertCircle,
  ReceiptText,
  Banknote,
  WalletCards,
  CalendarDays,
  History
} from 'lucide-react';
import './styles.css';

const categoryIcon = {
  Huevos: '🥚',
  'Miel y dulces': '🍯',
  'Frutos secos': '🥜',
  Aceites: '🫗',
  Conservas: '🫒',
  Snacks: '🍟',
  Almacén: '🧉',
  Artesanales: '🍪'
};

const money = n =>
  new Intl.NumberFormat('es-AR', {
    style: 'currency',
    currency: 'ARS',
    maximumFractionDigits: 0
  }).format(Number(n) || 0);

const todayISO = () => {
  const d = new Date();

  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(
    d.getDate()
  ).padStart(2, '0')}`;
};

const dateLabel = date =>
  new Intl.DateTimeFormat('es-AR', {
    weekday: 'long',
    day: '2-digit',
    month: '2-digit',
    year: 'numeric'
  }).format(new Date(`${date}T12:00:00`));

const dateShort = date =>
  new Intl.DateTimeFormat('es-AR', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric'
  }).format(new Date(`${date}T12:00:00`));

  const LOGIN_USERS = {
    bartol: 'angelbartol802@gmail.com',
    empleados: 'santquiroga10@gmail.com'
  };

function App() {
  const [products, setProducts] = useState([]);

  const [user, setUser] = useState(null);
  const [role, setRole] = useState(null);
  const [authLoading, setAuthLoading] = useState(true);

  const [loginOpen, setLoginOpen] = useState(false);
  const [logoutConfirm, setLogoutConfirm] = useState(false);
  const [loginForm, setLoginForm] = useState({
    username: '',
    password: ''
  });
  const [loginError, setLoginError] = useState('');
  const [loginLoading, setLoginLoading] = useState(false);

  const [sales, setSales] = useState([]);

  const [section, setSection] = useState('products');

  const [search, setSearch] = useState('');
  const [category, setCategory] = useState('Todos');
  const [sort, setSort] = useState('name-asc');
  const [viewMode, setViewMode] = useState('grid');

  const [modal, setModal] = useState(null);

  const [form, setForm] = useState({
    name: '',
    price: '',
    cost: '',
    category: 'Huevos'
  });

  const [newCategory, setNewCategory] = useState('');
  const [confirm, setConfirm] = useState(null);
  const [restoreConfirm, setRestoreConfirm] = useState(null);
  const [stockModal, setStockModal] = useState(null);
  const [stockValue, setStockValue] = useState('');
  const [notification, setNotification] = useState(null);

  const [saleDate, setSaleDate] = useState(todayISO());
  const [saleModal, setSaleModal] = useState(null);
  const [saleRegisterDate, setSaleRegisterDate] = useState('');

  const [saleForm, setSaleForm] = useState({
    description: '',
    paymentMethod: 'Efectivo',
    paymentStatus: 'pagada',
    amountPaid: ''
  });

  const [saleItems, setSaleItems] = useState([
    {
      productId: '',
      quantity: 1,
      unitPrice: ''
    }
  ]);

  const [saleItemSearches, setSaleItemSearches] = useState({});
  const [historyDate, setHistoryDate] = useState(todayISO());

  const [salesMonth, setSalesMonth] = useState(
    todayISO().slice(0, 7)
  );
  const [salesSearch, setSalesSearch] = useState('');

  const [salesView, setSalesView] = useState('day');

  const [paymentModal, setPaymentModal] = useState(null);
  const [paymentForm, setPaymentForm] = useState({
    date: todayISO(),
    amount: '',
    paymentMethod: 'Efectivo'
  });

  const [paymentHistoryModal, setPaymentHistoryModal] = useState(null);
  const [editingPayment, setEditingPayment] = useState(null);

  const isStaff = role === 'owner' || role === 'employee';
  const isOwner = role === 'owner';

  useEffect(() => {
    const loadSession = async () => {
      const {
        data: { session }
      } = await supabase.auth.getSession();

      if (!session) {
        setUser(null);
        setRole(null);
        setAuthLoading(false);
        return;
      }

      setUser(session.user);

      const { data: profile, error } = await supabase
        .from('perfiles')
        .select('rol')
        .eq('user_id', session.user.id)
        .single();

      if (error) {
        console.error('Error cargando perfil:', error);
        setRole(null);
      } else {
        setRole(profile.rol);
      }

      setAuthLoading(false);
    };

    loadSession();

    const {
      data: { subscription }
    } = supabase.auth.onAuthStateChange(() => {
      loadSession();
    });

    return () => {
      subscription.unsubscribe();
    };
  }, []);

  useEffect(() => {
    const loadProducts = async () => {
      const { data, error } = await supabase
        .from('productos')
        .select('id, nombre, precio_venta, categoria,activo')
        .order('id', { ascending: true });

      if (error) {
        console.error('Error cargando productos:', error);
        return;
      }

      let costs = [];

      if (role === 'owner') {
        const { data: costData, error: costError } = await supabase
          .from('producto_costos')
          .select('producto_id, precio_costo');

        if (costError) {
          console.error(
            'Error cargando costos de productos:',
            costError
          );
        } else {
          costs = costData || [];
        }
      }

      let stocks = [];
        if (isStaff) {
          const { data: stockData, error: stockError } = await supabase
            .from('producto_stock')
            .select('producto_id, stock');

          if (stockError) {
            console.error(
              'Error cargando stock de productos:',
              stockError
            );
          } else {
            stocks = stockData || [];
          }
        }

      const formattedProducts = data.map(product => {
        const productCost = costs.find(
          c => c.producto_id === product.id
        );
        const productStock = stocks.find(
          s => s.producto_id === product.id
        );

        return {
          id: product.id,
          name: product.nombre,
          price: Number(product.precio_venta),
          cost:
            role === 'owner' && productCost
              ? Number(productCost.precio_costo)
              : 0,
          stock:
            isStaff && productStock
              ? Number(productStock.stock)
              : 0,
          category: product.categoria,
          active: product.activo
        };
      });

      setProducts(formattedProducts);
    };

    loadProducts();
  }, [role, isStaff]);

  useEffect(() => {
    if (notification) {
      const t = setTimeout(() => setNotification(null), 2800);

      return () => clearTimeout(t);
    }
  }, [notification]);

  useEffect(() => {
    const loadSales = async () => {
      if (!isStaff) {
        setSales([]);
        return;
      }

      const { data, error } = await supabase
        .from('ventas')
        .select(`
          id,
          fecha,
          descripcion,
          total,
          medio_pago,
          registrado_por,
          created_at,

          venta_items (
            id,
            producto_id,
            producto_nombre,
            cantidad,
            precio_unitario,
            subtotal
          ),

          venta_pagos (
            id,
            fecha,
            monto,
            medio_pago,
            registrado_por,
            created_at
          )
        `)
        .order('fecha', { ascending: false })
        .order('created_at', { ascending: false });

      if (error) {
        console.error('Error cargando ventas:', error);
        return;
      }

      let finances = [];

      if (isOwner) {
        const {
          data: financeData,
          error: financeError
        } = await supabase
          .from('venta_item_finanzas')
          .select(`
            venta_item_id,
            costo_unitario,
            ganancia
          `);

        if (financeError) {
          console.error(
            'Error cargando finanzas de ventas:',
            financeError
          );
        } else {
          finances = financeData || [];
        }
      }

      const formattedSales = data.map(sale => {
        const items = (sale.venta_items || []).map(item => {
          const finance = finances.find(
            f => f.venta_item_id === item.id
          );

          return {
            id: item.id,
            productId: item.producto_id,
            productName: item.producto_nombre,
            quantity: Number(item.cantidad),
            unitPrice: Number(item.precio_unitario),
            subtotal: Number(item.subtotal),
            unitCost:
              isOwner && finance
                ? finance.costo_unitario !== null
                  ? Number(finance.costo_unitario)
                  : null
                : null,
            profit:
              isOwner && finance
                ? finance.ganancia !== null
                  ? Number(finance.ganancia)
                  : null
                : null
          };
        });

        const payments = (sale.venta_pagos || [])
          .map(payment => ({
            id: payment.id,
            date: payment.fecha,
            amount: Number(payment.monto),
            paymentMethod: payment.medio_pago,
            registeredBy: payment.registrado_por,
            createdAt: payment.created_at
          }))
          .sort((a, b) => {
            if (a.date !== b.date) {
              return a.date.localeCompare(b.date);
            }

            return new Date(a.createdAt) - new Date(b.createdAt);
          });

        const paid = payments.reduce(
          (sum, payment) => sum + payment.amount,
          0
        );

        const pending = Math.max(
          0,
          Number(sale.total) - paid
        );

        const paymentStatus =
          paid <= 0
            ? 'pendiente'
            : pending > 0
            ? 'parcial'
            : 'pagada';

        const hasPendingProfit =
          isOwner &&
          items.some(
            item =>
              item.unitCost === null ||
              item.profit === null
          );

        const totalProfit = isOwner
          ? hasPendingProfit
            ? null
            : items.reduce(
                (sum, item) => sum + item.profit,
                0
              )
          : null;

        return {
          id: sale.id,
          date: sale.fecha,
          description: sale.descripcion || '',
          total: Number(sale.total),

          paymentMethod: sale.medio_pago,

          payments,
          paid,
          pending,
          paymentStatus,

          registeredBy: sale.registrado_por,
          createdAt: sale.created_at,

          items,
          profit: totalProfit
        };
      });

      setSales(formattedSales);
    };

    loadSales();
  }, [isStaff, isOwner]);

  const categories = [
    'Todos',
    'Huevos',
    'Miel y dulces',
    'Frutos secos',
    'Aceites',
    'Conservas',
    'Snacks',
    'Almacén',
    'Artesanales',
    ...new Set(
      products
        .filter(p => p.active)
        .map(p => p.category)
    ),
    ...(isStaff ? ['Productos eliminados'] : [])
  ].filter((c, i, a) => a.indexOf(c) === i);

  const filtered = useMemo(() => {
    const list = products.filter(p => {
      const matchesSearch =
        p.name.toLowerCase().includes(search.toLowerCase());

      if (category === 'Productos eliminados') {
        return !p.active && matchesSearch;
      }

      if (!p.active) {
        return false;
      }

      return (
        (category === 'Todos' || p.category === category) &&
        matchesSearch
      );
    });

    return [...list].sort((a, b) =>
      sort === 'name-desc'
        ? b.name.localeCompare(a.name)
        : sort === 'price-asc'
        ? a.price - b.price
        : sort === 'price-desc'
        ? b.price - a.price
        : a.name.localeCompare(b.name)
    );
  }, [products, category, search, sort]);

  const daySales = useMemo(
    () => sales.filter(s => s.date === saleDate),
    [sales, saleDate]
  );

  const monthSales = useMemo(
    () =>
      sales
        .filter(s =>
          s.date.startsWith(salesMonth)
        )
        .sort((a, b) => {
          if (a.date !== b.date) {
            return a.date.localeCompare(b.date);
          }

          return new Date(b.createdAt) - new Date(a.createdAt);
        }),
    [sales, salesMonth]
  );

  const dayPayments = useMemo(
    () =>
      sales.flatMap(sale =>
        (sale.payments || []).filter(
          payment => payment.date === saleDate
        )
      ),
    [sales, saleDate]
  );

  const monthPayments = useMemo(
    () =>
      sales.flatMap(sale =>
        (sale.payments || []).filter(payment =>
          payment.date.startsWith(salesMonth)
        )
      ),
    [sales, salesMonth]
  );

  const monthTotals = useMemo(() => {
    const result = {
      total: 0,
      cash: 0,
      mp: 0,
      profit: 0,
      pendingPayments: false,
      pendingCost: false
    };

    monthSales.forEach(sale => {
      if (sale.paymentStatus !== 'pagada') {
        result.pendingPayments = true;
      }

      if (
        sale.profit !== null &&
        sale.profit !== undefined
      ) {
        result.profit += Number(sale.profit) || 0;
      } else {
        result.pendingCost = true;
      }
    });

    monthPayments.forEach(payment => {
      const amount = Number(payment.amount) || 0;

      result.total += amount;

      if (payment.paymentMethod === 'Mercado Pago') {
        result.mp += amount;
      } else {
        result.cash += amount;
      }
    });

    return result;
  }, [monthSales, monthPayments]);

  const searchedSales = useMemo(() => {
    const term = salesSearch
      .trim()
      .toLowerCase();

    if (!term) {
      return [];
    }

    return sales
      .filter(sale => {
        const descriptionMatch =
          sale.description
            ?.toLowerCase()
            .includes(term);

        const productMatch =
          sale.items.some(item =>
            item.productName
              ?.toLowerCase()
              .includes(term)
          );

        return descriptionMatch || productMatch;
      })
      .sort((a, b) => {
        if (a.date !== b.date) {
          return b.date.localeCompare(a.date);
        }

        return new Date(b.createdAt) - new Date(a.createdAt);
      });
  }, [sales, salesSearch]);

  const totals = useMemo(() => {
    const result = {
      total: 0,
      cash: 0,
      mp: 0,
      profit: 0,
      pendingPayments: false,
      pendingCost: false
    };

    daySales.forEach(sale => {
      if (sale.paymentStatus !== 'pagada') {
        result.pendingPayments = true;
      }

      if (
        sale.profit !== null &&
        sale.profit !== undefined
      ) {
        result.profit += Number(sale.profit) || 0;
      } else {
        result.pendingCost = true;
      }
    });

    dayPayments.forEach(payment => {
      const amount = Number(payment.amount) || 0;

      result.total += amount;

      if (payment.paymentMethod === 'Mercado Pago') {
        result.mp += amount;
      } else {
        result.cash += amount;
      }
    });

    return result;
  }, [daySales, dayPayments]);

  const showNotification = (message, type = 'success') =>
    setNotification({
      message,
      type
    });

  const login = async e => {
    e.preventDefault();
    setLoginError('');
    const username = loginForm.username
      .trim()
      .toLowerCase();
    const email = LOGIN_USERS[username];
    if (!email) {
      setLoginError('Usuario o contraseña incorrectos');
      return;
    }
    if (!loginForm.password) {
      setLoginError('Ingresá la contraseña');
      return;
    }
    setLoginLoading(true);
    const { error } = await supabase.auth.signInWithPassword({
      email,
      password: loginForm.password
    });
    setLoginLoading(false);
    if (error) {
      setLoginError('Usuario o contraseña incorrectos');
      return;
    }
    setLoginForm({
      username: '',
      password: ''
    });
    setLoginOpen(false);
    showNotification('Sesión iniciada');
  };

  const logout = async () => {
    const { error } = await supabase.auth.signOut();

    if (error) {
      console.error('Error cerrando sesión:', error);
      return;
    }

    setSection('products');

    showNotification('Sesión cerrada');
  };

  const openAdd = () => {
    setForm({
      name: '',
      price: '',
      cost: '',
      category: categories.find(c => c !== 'Todos') || 'Huevos'
    });

    setNewCategory('');
    setModal('add');
  };

  const openEdit = p => {
    setForm({
      name: p.name,
      price: p.price,
      cost: p.cost || '',
      category: p.category
    });
    setNewCategory('');
    setModal(p);
  };

  const openStockEdit = product => {
    setStockModal(product);
    setStockValue(String(product.stock ?? 0));
  };

  const saveStock = async e => {
    e.preventDefault();

    const newStock = Number(stockValue);

    if (!Number.isInteger(newStock) || newStock < 0) {
      showNotification(
        'El stock debe ser un número entero igual o mayor a 0',
        'warning'
      );
      return;
    }

    const { error } = await supabase
      .from('producto_stock')
      .update({
        stock: newStock,
        updated_at: new Date().toISOString()
      })
      .eq('producto_id', stockModal.id);

    if (error) {
      console.error('Error actualizando stock:', error);

      showNotification(
        'No se pudo actualizar el stock',
        'warning'
      );

      return;
    }

    setProducts(ps =>
      ps.map(p =>
        p.id === stockModal.id
          ? {
              ...p,
              stock: newStock
            }
          : p
      )
    );

    setStockModal(null);
    setStockValue('');

    showNotification('Stock actualizado');
  };

  const save = async e => {
    e.preventDefault();

    const name = form.name.trim();

    if (!name) return;

    const price = Math.max(0, Number(form.price) || 0);
    const cost = Math.max(0, Number(form.cost) || 0);

    const selectedCategory =
      form.category === '__new__'
        ? newCategory.trim()
        : form.category;

    if (!selectedCategory) return;

    if (modal === 'add') {
      const { data: newProduct, error } = await supabase
        .from('productos')
        .insert({
          nombre: name,
          precio_venta: price,
          categoria: selectedCategory
        })
        .select('id, nombre, precio_venta, categoria, activo')
        .single();

      if (error) {
        console.error('Error agregando producto:', error);
        showNotification('No se pudo agregar el producto', 'warning');
        return;
      }

      if (isOwner && cost > 0) {
        const { error: costError } = await supabase
          .from('producto_costos')
          .insert({
            producto_id: newProduct.id,
            precio_costo: cost
          });

        if (costError) {
          console.error(
            'Error guardando precio de costo:',
            costError
          );
        }
      }

      setProducts(ps => [
        ...ps,
        {
          id: newProduct.id,
          name: newProduct.nombre,
          price: Number(newProduct.precio_venta),
          cost: isOwner ? cost : 0,
          stock: 0,
          category: newProduct.categoria,
          active: newProduct.activo
        }
      ]);

      showNotification('Producto agregado');
    } else {
      const { data: updatedProduct, error } = await supabase
        .from('productos')
        .update({
          nombre: name,
          precio_venta: price,
          categoria: selectedCategory
        })
        .eq('id', modal.id)
        .select('id, nombre, precio_venta, categoria')
        .single();

      if (error) {
        console.error('Error actualizando producto:', error);
        showNotification(
          'No se pudo actualizar el producto',
          'warning'
        );
        return;
      }

      if (isOwner) {
        const { error: costError } = await supabase
          .from('producto_costos')
          .upsert({
            producto_id: modal.id,
            precio_costo: cost
          });

        if (costError) {
          console.error(
            'Error actualizando precio de costo:',
            costError
          );
          showNotification(
            'Producto actualizado, pero hubo un problema con el costo',
            'warning'
          );
        }
      }

      setProducts(ps =>
        ps.map(p =>
          p.id === modal.id
            ? {
                ...p,
                name: updatedProduct.nombre,
                price: Number(updatedProduct.precio_venta),
                cost: isOwner ? cost : p.cost,
                category: updatedProduct.categoria
              }
            : p
        )
      );

      showNotification('Producto actualizado');
    }

    setModal(null);
    setNewCategory('');
  };

  const remove = async () => {
    const { error } = await supabase
      .from('productos')
      .update({
        activo: false
      })
      .eq('id', confirm.id);

    if (error) {
      console.error('Error eliminando producto:', error);

      showNotification(
        'No se pudo eliminar el producto',
        'warning'
      );

      return;
    }

    setProducts(ps =>
      ps.map(p =>
        p.id === confirm.id
          ? {
              ...p,
              active: false
            }
          : p
      )
    );

    setConfirm(null);

    showNotification(
      'Producto movido a eliminados',
      'warning'
    );
  };

  const restoreProduct = async product => {
    const { error } = await supabase
      .from('productos')
      .update({
        activo: true
      })
      .eq('id', product.id);

    if (error) {
      console.error('Error restaurando producto:', error);

      showNotification(
        'No se pudo restaurar el producto',
        'warning'
      );

      return;
    }

    setProducts(ps =>
      ps.map(p =>
        p.id === product.id
          ? {
              ...p,
              active: true
            }
          : p
      )
    );

    showNotification('Producto restaurado');
  };

  const clearSearch = () => setSearch('');

  const changeSaleDate = days => {
    const date = new Date(`${saleDate}T12:00:00`);

    date.setDate(date.getDate() + days);

    const newDate = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(
      2,
      '0'
    )}-${String(date.getDate()).padStart(2, '0')}`;

    setSaleDate(newDate);
    setHistoryDate(newDate);
  };

  const changeSalesMonth = months => {
    const [year, month] = salesMonth
      .split('-')
      .map(Number);

    const date = new Date(
      year,
      month - 1 + months,
      1
    );

    const newMonth = `${date.getFullYear()}-${String(
      date.getMonth() + 1
    ).padStart(2, '0')}`;

    setSalesMonth(newMonth);
  };

  const salesMonthLabel = new Intl.DateTimeFormat(
    'es-AR',
    {
      month: 'long',
      year: 'numeric'
    }
  ).format(
    new Date(`${salesMonth}-01T12:00:00`)
  );

  const exportPrices = async () => {
    const exportProducts = products
      .filter(p => p.active)
      .sort((a, b) => {
        if (a.category !== b.category) {
          return a.category.localeCompare(b.category);
        }

        return a.name.localeCompare(b.name);
      });

    if (exportProducts.length === 0) {
      showNotification(
        'No hay productos para exportar',
        'warning'
      );
      return;
    }

    try {
      const workbook = new ExcelJS.Workbook();

      const worksheet = workbook.addWorksheet(
        'Lista de precios'
      );

      /*
        Ancho de columnas:
        A Producto
        B Precio de Venta
        C Categoría
      */
      worksheet.columns = [
        { width: 42 },
        { width: 20 },
        { width: 22 }
      ];

      /*
        Encabezados.
      */
      const headerRow = worksheet.addRow([
        'Producto',
        'Precio de Venta',
        'Categoría'
      ]);

      /*
        Color celeste igual al utilizado
        en los encabezados del Excel de ventas.
      */
      headerRow.eachCell(cell => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb: 'FFD9EAF7'
          }
        };

        cell.font = {
          bold: true
        };

        cell.border = {
          top: {
            style: 'thin',
            color: { argb: 'FFD9D9D9' }
          },
          bottom: {
            style: 'thin',
            color: { argb: 'FFD9D9D9' }
          },
          left: {
            style: 'thin',
            color: { argb: 'FFD9D9D9' }
          },
          right: {
            style: 'thin',
            color: { argb: 'FFD9D9D9' }
          }
        };

        cell.alignment = {
          vertical: 'middle'
        };
      });

      /*
        Productos.
      */
      exportProducts.forEach(product => {
        const price =
          Number(product.price) === 0
            ? 'PRECIO PENDIENTE'
            : Number(product.price);

        const row = worksheet.addRow([
          product.name,
          price,
          product.category
        ]);

        /*
          Formato moneda únicamente cuando
          el producto tiene un precio numérico.
        */
        const priceCell = row.getCell(2);

        if (typeof priceCell.value === 'number') {
          priceCell.numFmt = '$ #,##0';
        }
      });

      /*
        Permite filtrar y ordenar desde Excel.
      */
      worksheet.autoFilter = {
        from: 'A1',
        to: `C${worksheet.rowCount}`
      };

      /*
        Dejamos fija la fila de encabezados.
      */
      worksheet.views = [
        {
          state: 'frozen',
          xSplit: 0,
          ySplit: 1,
          topLeftCell: 'A2',
          activeCell: 'A2'
        }
      ];

      /*
        Generamos y descargamos el archivo.
      */
      const buffer =
        await workbook.xlsx.writeBuffer();

      const blob = new Blob(
        [buffer],
        {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }
      );

      const url =
        URL.createObjectURL(blob);

      const link =
        document.createElement('a');

      link.href = url;

      link.download =
        `lista-precios-bartol-${todayISO()}.xlsx`;

      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);

      URL.revokeObjectURL(url);

      showNotification(
        'Lista de precios exportada a Excel'
      );
    } catch (error) {
      console.error(
        'Error exportando lista de precios:',
        error
      );

      showNotification(
        'No se pudo exportar la lista de precios',
        'warning'
      );
    }
  };

  const openSaleAdd = () => {
    setSaleForm({
      description: '',
      paymentMethod: 'Efectivo',
      paymentStatus: 'pagada',
      amountPaid: ''
    });

    setSaleItems([
      {
        productId: '',
        quantity: 1,
        unitPrice: ''
      }
    ]);

    setSaleItemSearches({});

    setSaleRegisterDate(
      salesView === 'month'
        ? ''
        : saleDate
    );

    setSaleModal('add');
  };

  const openSaleEdit = s => {
    setSaleForm({
      description: s.description || '',
      paymentMethod: s.paymentMethod
    });

    setSaleItems(
      s.items.map(item => ({
        id: item.id,
        productId: item.productId,
        quantity: item.quantity,
        unitPrice: item.unitPrice
      }))
    );

    setSaleItemSearches({});

    setSaleModal(s);
  };

  const addSaleItem = () => {
    setSaleItems(items => [
      ...items,
      {
        productId: '',
        quantity: 1,
        unitPrice: ''
      }
    ]);
  };

  const removeSaleItem = index => {
    setSaleItems(items => {
      if (items.length === 1) {
        return items;
      }

      return items.filter((_, i) => i !== index);
    });
  };

  const updateSaleItem = (index, field, value) => {
    setSaleItems(items =>
      items.map((item, i) =>
        i === index
          ? {
              ...item,
              [field]: value
            }
          : item
      )
    );
  };

  const saleTotal = useMemo(
    () =>
      saleItems.reduce(
        (total, item) =>
          total +
          Number(item.quantity || 0) *
            Number(item.unitPrice || 0),
        0
      ),
    [saleItems]
  );

  const saveSale = async e => {
    e.preventDefault();

      if (!saleForm.description.trim()) {
        showNotification(
          'Ingresá una descripción para la venta',
          'warning'
        );
        return;
      }

    const validItems = saleItems.map(item => ({
      id: item.id,
      productId: item.productId,
      quantity: Number(item.quantity),
      unitPrice: Number(item.unitPrice)
    }));

    const invalidItem = saleItems.some(
      item =>
        !item.productId ||
        item.unitPrice === '' ||
        !Number.isInteger(Number(item.quantity)) ||
        Number(item.quantity) <= 0 ||
        !Number.isFinite(Number(item.unitPrice)) ||
        Number(item.unitPrice) < 0
    );

    if (invalidItem) {
      showNotification(
        'Revisá los productos, cantidades y precios',
        'warning'
      );
      return;
    }

    const productIds = validItems.map(item =>
      String(item.productId)
    );

    if (new Set(productIds).size !== productIds.length) {
      showNotification(
        'Un producto no puede aparecer dos veces en la misma venta',
        'warning'
      );
      return;
    }

    const itemsForSupabase = validItems.map(item => ({
      ...(item.id ? { id: item.id } : {}),
      producto_id: Number(item.productId),
      cantidad: item.quantity,
      precio_unitario: item.unitPrice
    }));

    let finalPaymentStatus = saleForm.paymentStatus;
    let finalAmountPaid =
      saleForm.paymentStatus === 'parcial'
        ? Number(saleForm.amountPaid)
        : null;

    if (
      saleModal === 'add' &&
      saleForm.paymentStatus === 'parcial'
    ) {
      if (
        !Number.isFinite(finalAmountPaid) ||
        finalAmountPaid <= 0
      ) {
        showNotification(
          'Ingresá un monto abonado mayor a 0',
          'warning'
        );
        return;
      }

      if (finalAmountPaid > saleTotal) {
        showNotification(
          `El monto abonado no puede superar el total de la venta (${money(
            saleTotal
          )})`,
          'warning'
        );
        return;
      }

      if (finalAmountPaid === saleTotal) {
        finalPaymentStatus = 'pagada';
        finalAmountPaid = null;
      }
    }

    let saleId;

    if (saleModal === 'add') {
      if (salesView === 'month' && !saleRegisterDate) {
        showNotification(
          'Seleccioná la fecha de la venta',
          'warning'
        );
        return;
      }
      /*
        En una venta nueva comprobamos el stock visible
        antes de enviarla.
      */
      for (const item of validItems) {
        const product = products.find(
          p => String(p.id) === String(item.productId)
        );

        if (!product) {
          showNotification(
            'Uno de los productos seleccionados no existe',
            'warning'
          );
          return;
        }

        if (item.quantity > Number(product.stock)) {
          showNotification(
            `Stock insuficiente de ${product.name}. Disponible: ${product.stock}`,
            'warning'
          );
          return;
        }
      }

      const { data, error } = await supabase.rpc(
        'registrar_venta_multiple',
        {
          p_fecha:
            salesView === 'month'
              ? saleRegisterDate
              : saleDate,

          p_descripcion: saleForm.description.trim(),

          p_medio_pago:
            saleForm.paymentStatus === 'pendiente'
              ? null
              : saleForm.paymentMethod,

          p_items: itemsForSupabase,

          p_estado_pago: finalPaymentStatus,

          p_monto_pagado:
            finalPaymentStatus === 'parcial'
              ? finalAmountPaid
              : finalPaymentStatus === 'pendiente'
              ? 0
              : null
        }
      );

      if (error) {
        console.error(
          'Error registrando venta múltiple:',
          error
        );

        showNotification(
          'No se pudo registrar la venta',
          'warning'
        );
        return;
      }

      saleId = data;
    } else {
      /*
        En edición dejamos que Supabase compruebe el stock,
        porque los triggers primero restauran las cantidades
        anteriores y luego descuentan las nuevas.
      */
      const { error } = await supabase.rpc(
        'actualizar_venta_multiple',
        {
          p_venta_id: saleModal.id,
          p_descripcion: saleForm.description.trim(),
          p_medio_pago: saleForm.paymentMethod,
          p_items: itemsForSupabase
        }
      );

      if (error) {
        console.error(error);

        if (
          error.message?.includes(
            'El total de la venta no puede quedar por debajo de lo ya pagado'
          )
        ) {
          showNotification(
            'El total de la venta no puede quedar por debajo de lo ya pagado.',
            'warning'
          );

          return;
        }

        if (
          error.message?.toLowerCase().includes('stock')
        ) {
          showNotification(
            'No se pudo actualizar la venta. Revisá el stock disponible.',
            'warning'
          );

          return;
        }

        showNotification(
          error.message || 'No se pudo actualizar la venta.',
          'warning'
        );

        return;
      }

      /*
        Si la venta tiene un único pago, permitimos
        corregir su medio de pago desde Editar venta.
      */
      if (saleModal.payments?.length === 1) {
        const payment = saleModal.payments[0];

        if (
          payment.paymentMethod !==
          saleForm.paymentMethod
        ) {
          const { error: paymentMethodError } =
            await supabase.rpc(
              'actualizar_medio_pago',
              {
                p_pago_id: payment.id,
                p_medio_pago:
                  saleForm.paymentMethod
              }
            );

          if (paymentMethodError) {
            console.error(
              'Error actualizando medio de pago:',
              paymentMethodError
            );

            showNotification(
              'La venta se actualizó, pero no se pudo modificar el medio de pago',
              'warning'
            );

            return;
          }
        }
      }

      saleId = saleModal.id;
    }

    /*
      Traemos nuevamente la venta desde Supabase porque
      los subtotales, total y ganancias fueron calculados
      por la base de datos.
    */
    const { data: saleData, error: saleLoadError } =
      await supabase
        .from('ventas')
        .select(`
          id,
          fecha,
          descripcion,
          total,
          medio_pago,
          registrado_por,
          created_at,

          venta_items (
            id,
            producto_id,
            producto_nombre,
            cantidad,
            precio_unitario,
            subtotal
          ),

          venta_pagos (
            id,
            fecha,
            monto,
            medio_pago,
            registrado_por,
            created_at
          )
        `)
        .eq('id', saleId)
        .single();

    if (saleLoadError) {
      console.error(
        'Error cargando la venta:',
        saleLoadError
      );

      showNotification(
        'La venta se guardó, pero no se pudo actualizar la pantalla',
        'warning'
      );
      return;
    }

    let finances = [];

    if (isOwner) {
      const itemIds = saleData.venta_items.map(
        item => item.id
      );

      if (itemIds.length > 0) {
        const {
          data: financeData,
          error: financeError
        } = await supabase
          .from('venta_item_finanzas')
          .select(`
            venta_item_id,
            costo_unitario,
            ganancia
          `)
          .in('venta_item_id', itemIds);

        if (financeError) {
          console.error(
            'Error cargando finanzas de la venta:',
            financeError
          );
        } else {
          finances = financeData || [];
        }
      }
    }

    const formattedItems = saleData.venta_items.map(item => {
      const finance = finances.find(
        f => f.venta_item_id === item.id
      );

      return {
        id: item.id,
        productId: item.producto_id,
        productName: item.producto_nombre,
        quantity: Number(item.cantidad),
        unitPrice: Number(item.precio_unitario),
        subtotal: Number(item.subtotal),
        unitCost:
          isOwner && finance
            ? finance.costo_unitario !== null
              ? Number(finance.costo_unitario)
              : null
            : null,
        profit:
          isOwner && finance
            ? finance.ganancia !== null
              ? Number(finance.ganancia)
              : null
            : null
      };
    });

    const payments = (saleData.venta_pagos || [])
      .map(payment => ({
        id: payment.id,
        date: payment.fecha,
        amount: Number(payment.monto),
        paymentMethod: payment.medio_pago,
        registeredBy: payment.registrado_por,
        createdAt: payment.created_at
      }))
      .sort((a, b) => {
        if (a.date !== b.date) {
          return a.date.localeCompare(b.date);
        }

        return new Date(a.createdAt) - new Date(b.createdAt);
      });

    const paid = payments.reduce(
      (sum, payment) => sum + payment.amount,
      0
    );

    const pending = Math.max(
      0,
      Number(saleData.total) - paid
    );

    const paymentStatus =
      paid <= 0
        ? 'pendiente'
        : pending > 0
        ? 'parcial'
        : 'pagada';

    const savedSale = {
      id: saleData.id,
      date: saleData.fecha,
      description: saleData.descripcion || '',
      total: Number(saleData.total),

      paymentMethod: saleData.medio_pago,

      payments,
      paid,
      pending,
      paymentStatus,

      registeredBy: saleData.registrado_por,
      createdAt: saleData.created_at,

      items: formattedItems,

      profit: isOwner
        ? formattedItems.some(
            item =>
              item.unitCost === null ||
              item.profit === null
          )
          ? null
          : formattedItems.reduce(
              (sum, item) => sum + item.profit,
              0
            )
        : null
    };

    if (saleModal === 'add') {
      setSales(ss => [
        savedSale,
        ...ss
      ]);
    } else {
      setSales(ss =>
        ss.map(s =>
          s.id === savedSale.id
            ? savedSale
            : s
        )
      );
    }

    /*
      Volvemos a leer el stock real de Supabase.
      En una edición es más seguro que calcularlo
      manualmente porque pueden agregarse, quitarse
      o cambiarse productos.
    */
    const {
      data: stockData,
      error: stockError
    } = await supabase
      .from('producto_stock')
      .select('producto_id, stock');

    if (stockError) {
      console.error(
        'Error actualizando stock en pantalla:',
        stockError
      );
    } else {
      setProducts(ps =>
        ps.map(product => {
          const stockRow = stockData.find(
            row => row.producto_id === product.id
          );

          return stockRow
            ? {
                ...product,
                stock: Number(stockRow.stock)
              }
            : product;
        })
      );
    }

    showNotification(
      saleModal === 'add'
        ? 'Venta registrada'
        : 'Venta actualizada'
    );

    setSaleModal(null);

    setSaleForm({
      description: '',
      paymentMethod: 'Efectivo',
      paymentStatus: 'pagada',
      amountPaid: ''
    });

    setSaleItems([
      {
        productId: '',
        quantity: 1,
        unitPrice: ''
      }
    ]);

    setSaleItemSearches({});
  };

  const openPaymentModal = sale => {
    setPaymentModal(sale);

    setPaymentForm({
      date: todayISO(),
      amount: '',
      paymentMethod: 'Efectivo'
    });
  };

  const savePayment = async e => {
    e.preventDefault();

    if (!paymentModal) return;

    const amount = Number(paymentForm.amount);

    if (!Number.isFinite(amount) || amount <= 0) {
      showNotification(
        'Ingresá un monto válido',
        'warning'
      );
      return;
    }

    if (amount > Number(paymentModal.pending)) {
      showNotification(
        'El pago no puede superar el saldo pendiente',
        'warning'
      );
      return;
    }

    const { error } = await supabase.rpc(
      'registrar_pago_venta',
      {
        p_venta_id: paymentModal.id,
        p_fecha: paymentForm.date,
        p_monto: amount,
        p_medio_pago: paymentForm.paymentMethod
      }
    );

    if (error) {
      console.error('Error registrando pago:', error);

      showNotification(
        error.message || 'No se pudo registrar el pago',
        'warning'
      );

      return;
    }

    const { data: paymentsData, error: paymentsError } =
      await supabase
        .from('venta_pagos')
        .select(`
          id,
          fecha,
          monto,
          medio_pago,
          registrado_por,
          created_at
        `)
        .eq('venta_id', paymentModal.id)
        .order('fecha', { ascending: true })
        .order('created_at', { ascending: true });

    if (paymentsError) {
      console.error(
        'Error actualizando pagos:',
        paymentsError
      );

      showNotification(
        'El pago se registró, pero no se pudo actualizar la pantalla',
        'warning'
      );

      return;
    }

    const payments = (paymentsData || []).map(payment => ({
      id: payment.id,
      date: payment.fecha,
      amount: Number(payment.monto),
      paymentMethod: payment.medio_pago,
      registeredBy: payment.registrado_por,
      createdAt: payment.created_at
    }));

    setSales(currentSales =>
      currentSales.map(sale => {
        if (sale.id !== paymentModal.id) {
          return sale;
        }

        const paid = payments.reduce(
          (sum, payment) => sum + payment.amount,
          0
        );

        const pending = Math.max(
          0,
          Number(sale.total) - paid
        );

        return {
          ...sale,
          payments,
          paid,
          pending,
          paymentStatus:
            paid <= 0
              ? 'pendiente'
              : pending > 0
              ? 'parcial'
              : 'pagada'
        };
      })
    );

    setPaymentModal(null);

    setPaymentForm({
      date: todayISO(),
      amount: '',
      paymentMethod: 'Efectivo'
    });

    showNotification('Pago registrado');
  };

  const savePaymentMethod = async () => {
    if (!editingPayment) return;

    const { error } = await supabase.rpc(
      'actualizar_medio_pago',
      {
        p_pago_id: editingPayment.id,
        p_medio_pago: editingPayment.paymentMethod
      }
    );

    if (error) {
      console.error(
        'Error actualizando medio de pago:',
        error
      );

      showNotification(
        'No se pudo actualizar el medio de pago',
        'warning'
      );

      return;
    }

    /*
      Actualizamos la venta en el estado principal.
    */
    setSales(currentSales =>
      currentSales.map(sale => {
        if (sale.id !== editingPayment.saleId) {
          return sale;
        }

        const updatedPayments = sale.payments.map(payment =>
          payment.id === editingPayment.id
            ? {
                ...payment,
                paymentMethod: editingPayment.paymentMethod
              }
            : payment
        );

        return {
          ...sale,
          payments: updatedPayments
        };
      })
    );

    /*
      Actualizamos también el historial que está abierto.
    */
    setPaymentHistoryModal(currentSale => {
      if (!currentSale) return null;

      return {
        ...currentSale,
        payments: currentSale.payments.map(payment =>
          payment.id === editingPayment.id
            ? {
                ...payment,
                paymentMethod: editingPayment.paymentMethod
              }
            : payment
        )
      };
    });

    setEditingPayment(null);

    showNotification('Medio de pago actualizado');
  };


  const removeSale = s => {
    setConfirm({
      type: 'sale',
      ...s
    });
  };

  const removeConfirmed = async () => {
    if (confirm.type === 'sale') {
      const saleToDelete = confirm;

      const { error } = await supabase
        .from('ventas')
        .delete()
        .eq('id', saleToDelete.id);

      if (error) {
        console.error('Error eliminando venta:', error);

        showNotification(
          'No se pudo eliminar la venta',
          'warning'
        );

        return;
      }

      setSales(ss =>
        ss.filter(s => s.id !== saleToDelete.id)
      );

      setProducts(ps =>
        ps.map(product => {
          const soldItem = saleToDelete.items.find(
            item =>
              String(item.productId) ===
              String(product.id)
          );

          if (!soldItem) {
            return product;
          }

          return {
            ...product,
            stock:
              Number(product.stock) +
              Number(soldItem.quantity)
          };
        })
      );

      setConfirm(null);

      showNotification(
        'Venta eliminada y stock restaurado',
        'warning'
      );

      return;
    }

    await remove();
  };

  const exportSalesExcelJS = async () => {
    if (!isOwner) {
      showNotification(
        'Solo el dueño puede exportar las ventas',
        'warning'
      );
      return;
    }

    if (sales.length === 0) {
      showNotification(
        'No hay ventas registradas para exportar',
        'warning'
      );
      return;
    }

    try {
      const workbook = new ExcelJS.Workbook();

      /*
        Agrupamos cada venta según el mes
        en que fue realizada.
      */
      const salesByMonth = {};

      sales.forEach(sale => {
        const monthKey = sale.date.slice(0, 7);

        if (!salesByMonth[monthKey]) {
          salesByMonth[monthKey] = [];
        }

        salesByMonth[monthKey].push(sale);
      });

      const sortedMonths = Object.keys(
        salesByMonth
      ).sort();

      sortedMonths.forEach(monthKey => {
        const monthSales = [
          ...salesByMonth[monthKey]
        ].sort((a, b) => {
          if (a.date !== b.date) {
            return a.date.localeCompare(b.date);
          }

          return (
            new Date(a.createdAt) -
            new Date(b.createdAt)
          );
        });

        /*
          Nombre de la hoja.
        */
        const [year, month] = monthKey.split('-');

        const monthName =
          new Intl.DateTimeFormat(
            'es-AR',
            { month: 'long' }
          ).format(
            new Date(
              Number(year),
              Number(month) - 1,
              1
            )
          );

        const rawSheetName =
          `${monthName} ${year}`;

        const sheetName =
          rawSheetName.charAt(0).toUpperCase() +
          rawSheetName.slice(1);

        const worksheet =
          workbook.addWorksheet(sheetName);

        /*
          Ancho de columnas.
        */
        worksheet.columns = [
          { width: 12 }, // N° Venta
          { width: 22 }, // Fecha
          { width: 33 }, // Producto
          { width: 12 }, // Cantidad
          { width: 17 }, // Precio Unitario
          { width: 12 }, // Subtotal
          { width: 17 }, // Costo Unitario
          { width: 12 }, // Ganancia
          { width: 26 }, // Descripción
          { width: 15 }, // Estado
          { width: 15 }, // Total Venta
          { width: 15 }, // Pagado
          { width: 15 }, // Pendiente
          { width: 22 }  // Medio de Pago
        ];

        /*
          FILA 1 - VENTAS DEL MES
        */
        worksheet.addRow([
          'VENTAS DEL MES'
        ]);

        worksheet.mergeCells('A1:N1');

        const titleCell = worksheet.getCell('A1');

        titleCell.font = {
          bold: true,
          color: { argb: 'FFFFFFFF' },
          size: 12
        };

        titleCell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FF1F4E78' }
        };

        titleCell.alignment = {
          vertical: 'middle',
          horizontal: 'center'
        };

        /*
          FILA 2 - ENCABEZADOS
        */
        const headerRow = worksheet.addRow([
          'N° Venta',
          'Fecha',
          'Producto',
          'Cantidad',
          'Precio Unitario',
          'Subtotal',
          'Costo Unitario',
          'Ganancia',
          'Descripción',
          'Estado',
          'Total Venta',
          'Pagado',
          'Pendiente',
          'Medio de Pago'
        ]);

        const thinBorder = {
          top: {
            style: 'thin',
            color: { argb: 'FFD9D9D9' }
          },
          bottom: {
            style: 'thin',
            color: { argb: 'FFD9D9D9' }
          },
          left: {
            style: 'thin',
            color: { argb: 'FFD9D9D9' }
          },
          right: {
            style: 'thin',
            color: { argb: 'FFD9D9D9' }
          }
        };

        headerRow.eachCell(cell => {
          cell.font = {
            bold: true
          };

          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFD9EAF7' }
          };

          cell.border = thinBorder;

          cell.alignment = {
            vertical: 'middle'
          };
        });

        /*
          Congelamos las dos primeras filas.
        */
        worksheet.views = [
          {
            state: 'frozen',
            xSplit: 0,
            ySplit: 2,
            topLeftCell: 'A3',
            activeCell: 'A3'
          }
        ];

        let totalSold = 0;
        let totalPaid = 0;
        let totalPending = 0;

        let totalCash = 0;
        let totalMercadoPago = 0;

        let totalProfit = 0;

        let hasPendingPayments = false;
        let hasPendingCost = false;

        /*
          Agregamos las ventas.
        */
        monthSales.forEach(sale => {
          const saleTotal =
            Number(sale.total) || 0;

          const salePaid =
            Number(sale.paid) || 0;

          const salePending =
            Number(sale.pending) || 0;

          totalSold += saleTotal;
          totalPaid += salePaid;
          totalPending += salePending;

          if (sale.paymentStatus !== 'pagada') {
            hasPendingPayments = true;
          }

          /*
            Totales según medio de pago.
          */
          (sale.payments || []).forEach(payment => {
            const amount =
              Number(payment.amount) || 0;

            if (
              payment.paymentMethod ===
              'Mercado Pago'
            ) {
              totalMercadoPago += amount;
            } else if (
              payment.paymentMethod ===
              'Efectivo'
            ) {
              totalCash += amount;
            }
          });

          const paymentMethods = [
            ...new Set(
              (sale.payments || [])
                .map(payment => payment.paymentMethod)
                .filter(Boolean)
            )
          ];

          let paymentMethodText = 'SIN PAGOS';

          if (paymentMethods.length > 0) {
            paymentMethodText =
              paymentMethods.join(' / ');
          }

          const statusText =
            sale.paymentStatus === 'pagada'
              ? 'PAGADA'
              : sale.paymentStatus === 'parcial'
              ? 'PAGO PARCIAL'
              : 'PENDIENTE';

          /*
            Color correspondiente al estado.
          */
          let backgroundColor;

          if (sale.paymentStatus === 'pagada') {
            backgroundColor = 'FFC6E0B4';
          } else if (
            sale.paymentStatus === 'parcial'
          ) {
            backgroundColor = 'FFFFE699';
          } else {
            backgroundColor = 'FFF4B7B2';
          }

          /*
            Una fila por producto.
          */
          (sale.items || []).forEach(
            (item, itemIndex) => {
              const itemProfit =
                item.profit !== null &&
                item.profit !== undefined
                  ? Number(item.profit)
                  : null;

              if (itemProfit === null) {
                hasPendingCost = true;
              } else {
                totalProfit += itemProfit;
              }

              const isFirstItem =
                itemIndex === 0;

              const row = worksheet.addRow([
                sale.id,

                isFirstItem
                  ? dateShort(sale.date)
                  : '',

                item.productName,

                Number(item.quantity),

                Number(item.unitPrice),

                Number(item.subtotal),

                item.unitCost !== null &&
                item.unitCost !== undefined
                  ? Number(item.unitCost)
                  : 'PENDIENTE',

                itemProfit !== null
                  ? itemProfit
                  : 'PENDIENTE',

                isFirstItem
                  ? sale.description || ''
                  : '',

                isFirstItem
                  ? statusText
                  : '',

                isFirstItem
                  ? saleTotal
                  : '',

                isFirstItem
                  ? salePaid
                  : '',

                isFirstItem
                  ? salePending
                  : '',

                isFirstItem
                  ? paymentMethodText
                  : ''
              ]);

              /*
                Color y bordes de toda la fila.
              */
              row.eachCell(
                { includeEmpty: true },
                cell => {
                  cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: {
                      argb: backgroundColor
                    }
                  };

                  cell.border = thinBorder;

                  cell.alignment = {
                    vertical: 'middle'
                  };
                }
              );

              /*
                Formato moneda.
                E, F, G, H? No:
                E Precio Unitario
                F Subtotal
                G Costo Unitario
                H Ganancia
                K Total Venta
                L Pagado
                M Pendiente
              */
              [
                5, 6, 7, 8,
                11, 12, 13
              ].forEach(column => {
                const cell =
                  row.getCell(column);

                if (
                  typeof cell.value === 'number'
                ) {
                  cell.numFmt = '$ #,##0';
                }
              });

              /*
                Última fila de la venta:
                borde inferior más marcado.
              */
              if (
                itemIndex ===
                (sale.items || []).length - 1
              ) {
                row.eachCell(
                  { includeEmpty: true },
                  cell => {
                    cell.border = {
                      ...(cell.border || {}),

                      bottom: {
                        style: 'thick',
                        color: {
                          argb: 'FF595959'
                        }
                      }
                    };
                  }
                );
              }
            }
          );
        });

        /*
          Última fila perteneciente a la tabla
          de ventas.
        */
        const salesTableLastRow =
          worksheet.rowCount;

        /*
          Filtro únicamente sobre las ventas.
        */
        if (salesTableLastRow >= 2) {
          worksheet.autoFilter = {
            from: 'A2',
            to: `N${salesTableLastRow}`
          };
        }

        /*
          Dos filas vacías.
        */
        worksheet.addRow([]);
        worksheet.addRow([]);

        /*
          RESUMEN DEL MES.
        */
        const summaryTitleRow =
          worksheet.addRow([
            'RESUMEN DEL MES'
          ]);

        worksheet.mergeCells(
          `A${summaryTitleRow.number}:N${summaryTitleRow.number}`
        );

        const summaryTitleCell =
          summaryTitleRow.getCell(1);

        summaryTitleCell.font = {
          bold: true,
          color: { argb: 'FFFFFFFF' },
          size: 12
        };

        summaryTitleCell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FF1F4E78' }
        };

        /*
          Calculamos resultado de ganancias.
        */
        let profitResult;

        if (
          hasPendingPayments &&
          hasPendingCost
        ) {
          profitResult =
            'HAY PAGOS PENDIENTES Y COSTOS PENDIENTES';
        } else if (hasPendingPayments) {
          profitResult =
            'HAY PAGOS PENDIENTES';
        } else if (hasPendingCost) {
          profitResult =
            'HAY COSTOS PENDIENTES';
        } else {
          profitResult = totalProfit;
        }

        const summaryData = [
          ['TOTAL VENDIDO', totalSold],
          ['TOTAL COBRADO', totalPaid],
          ['TOTAL PENDIENTE', totalPending],
          [
            'TOTAL MERCADO PAGO',
            totalMercadoPago
          ],
          ['TOTAL EFECTIVO', totalCash],
          ['TOTAL GANANCIAS', profitResult]
        ];

        summaryData.forEach(
          ([label, value]) => {
            const row = worksheet.addRow([
              '',
              label,
              '',
              '',
              value
            ]);

            const labelCell =
              row.getCell(2);

            const valueCell =
              row.getCell(5);

            labelCell.font = {
              bold: true
            };

            valueCell.font = {
              bold: true
            };

            labelCell.border =
              thinBorder;

            valueCell.border =
              thinBorder;

            if (
              typeof value === 'number'
            ) {
              valueCell.numFmt =
                '$ #,##0';
            }
          }
        );
      });

      /*
        Generamos el archivo.
      */
      const buffer =
        await workbook.xlsx.writeBuffer();

      const blob = new Blob(
        [buffer],
        {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }
      );

      const url =
        URL.createObjectURL(blob);

      const link =
        document.createElement('a');

      link.href = url;

      link.download =
        `ventas-bartol-${todayISO()}.xlsx`

      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);

      URL.revokeObjectURL(url);

      showNotification(
        'Todas las ventas fueron exportadas a Excel'
      );
    } catch (error) {
      console.error(
        'Error exportando ventas con ExcelJS:',
        error
      );

      showNotification(
        'No se pudieron exportar las ventas',
        'warning'
      );
    }
  };

  return (
    <div className="app">
      <div className="floating">🥚</div>
      <div className="floating f2">🍯</div>
      <div className="floating f3">🥜</div>
      <div className="grain" />

      <header>
        <div className="brand brand-image">
          <img
            src={bartolLogo}
            alt="Bartol - Reparto y distribución de huevos"
            className="bartol-logo"
          />
        </div>

        <div className={`header-actions ${
          section === 'products' && role === 'employee'
          ? 'employee-products-actions'
          : ''
        }`}
        >
          {/* EXPORTACIONES */}
          {section === 'products' && (
            <button
              className="secondary-action"
              onClick={exportPrices}
            >
              <Download size={17} />
              Exportar Precios
            </button>
          )}

          {section === 'sales' && isOwner && (
            <button
              className="secondary-action"
              onClick={exportSalesExcelJS}
            >
              <Download size={17} />
              Exportar Ventas
            </button>
          )}

          {/* SESIÓN */}
          {authLoading ? null : user ? (
            <div className="session-box">
              <button
                className="secondary-action"
                onClick={() => setLogoutConfirm(true)}
              >
                Cerrar sesión
              </button>

              <span>
                {role === 'owner'
                  ? 'Dueño'
                  : role === 'employee'
                  ? 'Empleado'
                  : 'Usuario'}
              </span>
            </div>
          ) : (
            <button
              className="secondary-action"
              onClick={() => {
                setLoginError('');
                setLoginOpen(true);
              }}
            >
              Iniciar sesión
            </button>
          )}
        </div>
      </header>

      <main>
        <nav className="section-switch">
          <button
            className={section === 'products' ? 'active' : ''}
            onClick={() => setSection('products')}
          >
            <Package size={18} />
            Productos
          </button>

          {isStaff && (
            <button
              className={section === 'sales' ? 'active' : ''}
              onClick={() => setSection('sales')}
            >
              <ReceiptText size={18} />
              Ventas
            </button>
          )}
        </nav>

        {section === 'products' ? (
          <>
            <section className="hero">
              <div className="hero-copy">
                <div className="eyebrow">
                  <Sparkles size={15} />
                  LISTA DE PRECIOS
                </div>

                <h1>
                  Todo Bartol,
                  <br />
                  <em>en un solo lugar.</em>
                </h1>

                <p>
                  Buscá, consultá y actualizá tus productos en segundos.
                </p>
              </div>

              <div className="hero-mascot">
                <img
                  src={huevoSkate}
                  alt=""
                  className="hero-skate-egg"
                />
              </div>
            </section>

            <section className="toolbar">
              <div className="search">
                <Search size={20} />

                <input
                  value={search}
                  onChange={e => setSearch(e.target.value)}
                  placeholder="Buscar productos..."
                />

                <button
                  className="clear-search"
                  onClick={clearSearch}
                  title="Limpiar búsqueda"
                  disabled={!search}
                >
                  <X size={16} />
                </button>
              </div>

              <div className="categories">
                {categories.map(c => (
                  <button
                    key={c}
                    className={category === c ? 'active' : ''}
                    onClick={() => setCategory(c)}
                  >
                    {c !== 'Todos' && categoryIcon[c]} {c}
                  </button>
                ))}
              </div>
            </section>

            <div className="summary">
              <div className="summary-left">
                {isStaff && (
                  <button
                    className="add"
                    onClick={openAdd}
                  >
                    <Plus size={19} />
                    Nuevo producto
                  </button>
                )}

                <span>
                  <Package size={17} />
                  <b>{filtered.length}</b> productos
                </span>
              </div>

              <div className="view-tools">
                <select
                  value={sort}
                  onChange={e => setSort(e.target.value)}
                >
                  <option value="name-asc">
                    Nombre A → Z
                  </option>

                  <option value="name-desc">
                    Nombre Z → A
                  </option>

                  <option value="price-asc">
                    Precio menor → mayor
                  </option>

                  <option value="price-desc">
                    Precio mayor → menor
                  </option>
                </select>

                <div className="view-switch">
                  <button
                    className={
                      viewMode === 'grid'
                        ? 'active'
                        : ''
                    }
                    onClick={() => setViewMode('grid')}
                    title="Vista grilla"
                  >
                    <LayoutGrid size={16} />
                  </button>

                  <button
                    className={
                      viewMode === 'list'
                        ? 'active'
                        : ''
                    }
                    onClick={() => setViewMode('list')}
                    title="Vista lista"
                  >
                    <List size={16} />
                  </button>
                </div>
              </div>
            </div>

            {filtered.length === 0 ? (
              <div className="empty">
                <div>🔎</div>
                <h2>No encontramos productos</h2>
                <p>Probá con otro nombre o categoría.</p>
              </div>
            ) : (
              <div
                className={
                  viewMode === 'grid'
                    ? 'grid'
                    : 'list-view'
                }
              >
                {filtered.map(p => (
                  <article
                    className="card"
                    key={p.id}
                  >
                    {isStaff && (
                      <div className="actions">
                        {p.active ? (
                          <>
                            <button
                              title="Editar"
                              onClick={() => openEdit(p)}
                            >
                              <Pencil size={16} />
                            </button>

                            <button
                              title="Editar stock"
                              onClick={() => openStockEdit(p)}
                            >
                              <Package size={16} />
                            </button>

                            <button
                              title="Eliminar"
                              onClick={() => setConfirm(p)}
                            >
                              <Trash2 size={16} />
                            </button>
                          </>
                        ) : (
                          <button
                            title="Restaurar producto"
                            onClick={() => setRestoreConfirm(p)}
                          >
                            <Package size={16} />
                          </button>
                        )}
                      </div>
                    )}

                    <div className="product-info">
                      <h2>{p.name}</h2>

                      <div className="product-prices">
                        <div
                          className={
                            p.price === 0
                              ? 'price pending'
                              : 'price'
                          }
                        >
                          {p.price === 0 ? (
                            <>
                              <AlertCircle size={14} />
                              Precio pendiente
                            </>
                          ) : (
                            money(p.price)
                          )}
                        </div>

                        {isOwner && (
                          <div className="admin-product-data">
                            <span>
                              Costo:{' '}
                              {p.cost > 0
                                ? money(p.cost)
                                : 'Pendiente'}
                            </span>

                            <span>
                              Ganancia:{' '}
                              {p.cost > 0 && p.price > 0
                                ? money(p.price - p.cost)
                                : 'Pendiente'}
                            </span>
                          </div>
                        )}
                        {isStaff && (
                          <div className="admin-product-data">
                            <span>
                              Stock: {p.stock}
                            </span>
                          </div>
                        )}
                      </div>
                    </div>

                    <span className="tag">
                      {categoryIcon[p.category] || '📦'}{' '}
                      {p.category}
                    </span>
                  </article>
                ))}
              </div>
            )}
          </>
        ) : (
          <>
            <section className="sales-summary">
              <div className="stat total">
                <span>
                  Total vendido del {salesView === 'day' ? 'día' : 'mes'}
                </span>

                <strong>
                  {money(
                    salesView === 'day'
                      ? totals.total
                      : monthTotals.total
                  )}
                </strong>

                <small>
                  {salesView === 'day'
                    ? daySales.length
                    : monthSales.length}{' '}
                  {(salesView === 'day'
                    ? daySales.length
                    : monthSales.length) === 1
                    ? 'venta'
                    : 'ventas'}
                </small>
              </div>

              <div className="stat">
                <span>
                  <Banknote size={16} />
                  Efectivo
                </span>

                <strong>
                  {money(
                    salesView === 'day'
                      ? totals.cash
                      : monthTotals.cash
                  )}
                </strong>
              </div>

              <div className="stat">
                <span>
                  <WalletCards size={16} />
                  Mercado Pago
                </span>

                <strong>
                  {money(
                    salesView === 'day'
                      ? totals.mp
                      : monthTotals.mp
                  )}
                </strong>
              </div>

              {isOwner && (
                <div className="stat profit-stat">
                  <span>
                    Ganancia del {salesView === 'day' ? 'día' : 'mes'}
                  </span>

                  {(
                    salesView === 'day'
                      ? totals.pendingPayments || totals.pendingCost
                      : monthTotals.pendingPayments ||
                        monthTotals.pendingCost
                  ) ? (
                    <>
                      <strong>—</strong>

                      <small>
                        {(
                          salesView === 'day'
                            ? totals.pendingPayments
                            : monthTotals.pendingPayments
                        ) &&
                        (
                          salesView === 'day'
                            ? totals.pendingCost
                            : monthTotals.pendingCost
                        )
                          ? 'Hay pagos pendientes y ventas con costo pendiente'
                          : (
                              salesView === 'day'
                                ? totals.pendingPayments
                                : monthTotals.pendingPayments
                            )
                          ? 'Hay pagos pendientes'
                          : 'Hay ventas con costo pendiente'}
                      </small>
                    </>
                  ) : (
                    <strong>
                      {money(
                        salesView === 'day'
                          ? totals.profit
                          : monthTotals.profit
                      )}
                    </strong>
                  )}
                </div>
              )}
            </section>

            
            {!salesSearch.trim() && (
              <div className="sales-view-switch">
                <button
                  className={salesView === 'day' ? 'active' : ''}
                  onClick={() => setSalesView('day')}
                >
                  Día
                </button>

                <button
                  className={salesView === 'month' ? 'active' : ''}
                  onClick={() => setSalesView('month')}
                >
                  Mes
                </button>
              </div>
            )}

            <div className="sales-search">
              <Search size={18} />

              <input
                type="text"
                placeholder="Buscar venta por producto o descripción..."
                value={salesSearch}
                onChange={e => setSalesSearch(e.target.value)}
              />

              {salesSearch && (
                <button
                  type="button"
                  className="sales-search-clear"
                  onClick={() => setSalesSearch('')}
                  title="Limpiar búsqueda"
                >
                  <X size={16} />
                </button>
              )}
            </div>

            <div className="sales-title">
              {salesSearch.trim() ? (
                <div className="sales-search-results-title">
                  <h2>
                    Resultados para "{salesSearch.trim()}"
                  </h2>

                  <span>
                    {searchedSales.length}{' '}
                    {searchedSales.length === 1
                      ? 'venta encontrada'
                      : 'ventas encontradas'}
                  </span>
                </div>
              ) : (
                <div className="sale-date-navigation">
                  <button
                    className="date-arrow"
                    onClick={() =>
                      salesView === 'day'
                        ? changeSaleDate(-1)
                        : changeSalesMonth(-1)
                    }
                    title={
                      salesView === 'day'
                        ? 'Día anterior'
                        : 'Mes anterior'
                    }
                  >
                    <ChevronLeft size={20} />
                  </button>

                  <h2 className="sales-month-title">
                    {salesView === 'day'
                      ? dateLabel(saleDate)
                      : salesMonthLabel}
                  </h2>

                  <button
                    className="date-arrow"
                    onClick={() =>
                      salesView === 'day'
                        ? changeSaleDate(1)
                        : changeSalesMonth(1)
                    }
                    title={
                      salesView === 'day'
                        ? 'Día siguiente'
                        : 'Mes siguiente'
                    }
                  >
                    <ChevronRight size={20} />
                  </button>
                </div>
              )}

              <button
                className="sale-add"
                onClick={openSaleAdd}
              >
                <Plus size={17} />
                Nueva venta
              </button>
            </div>

            {(salesSearch.trim()
              ? searchedSales
              : salesView === 'day'
                ? daySales
                : monthSales
            ).length === 0 ? (
              <div className="empty sales-empty">
                <div>🧾</div>

                <h2>
                  {salesSearch.trim()
                    ? 'Sin resultados'
                    : 'Todavía no hay ventas'}
                </h2>

                <p>
                  {salesSearch.trim()
                    ? `No encontramos ventas que coincidan con "${salesSearch.trim()}".`
                    : salesView === 'day'
                      ? 'Registrá la primera venta de este día para empezar a llevar el control.'
                      : 'No hay ventas registradas para este mes.'}
                </p>

                {!salesSearch.trim() && (
                  <button
                    className="sale-add"
                    onClick={openSaleAdd}
                  >
                    <Plus size={17} />
                    Registrar venta
                  </button>
                )}
              </div>
            ) : (
              <div className="sales-list">
                {(salesSearch.trim()
                  ? searchedSales
                  : salesView === 'day'
                    ? daySales
                    : monthSales
                ).map(s => (
                  <article
                    className="sale-card sale-card-multiple"
                    key={s.id}
                  >
                    <div className="sale-main sale-main-multiple">
                      <div className="sale-products">
                        {(salesView === 'month' || salesSearch.trim()) && (
                          <div className="sale-card-date">
                            <CalendarDays size={14} />
                            {dateShort(s.date)}
                          </div>
                        )}
                        {s.items.map(item => (
                          <div
                            className="sale-product-row"
                            key={item.id}
                          >
                            <div>
                              <h3>{item.productName}</h3>

                              <p>
                                {item.quantity} × {money(item.unitPrice)} c/u
                              </p>
                            </div>

                            {s.items.length > 1 && (
                              <strong>
                                {money(item.subtotal)}
                              </strong>
                            )}
                          </div>
                        ))}

                        {s.description && (
                          <p className="sale-description">
                            {s.description}
                          </p>
                        )}
                      </div>
                    </div>

                    <div className="sale-payment">
                      <span
                        className={`payment-status ${s.paymentStatus}`}
                      >
                        {s.paymentStatus === 'pagada'
                          ? 'PAGADA'
                          : s.paymentStatus === 'parcial'
                            ? 'PAGO PARCIAL'
                            : 'PENDIENTE'}
                      </span>

                      {s.paymentStatus !== 'pagada' ? (
                        <div className="sale-payment-summary">
                          <div>
                            <small>Total</small>
                            <strong>{money(s.total)}</strong>
                          </div>

                          <div>
                            <small>Pagado</small>
                            <strong>{money(s.paid)}</strong>
                          </div>

                          <div>
                            <small>Pendiente</small>
                            <strong className="pending-amount">
                              {money(s.pending)}
                            </strong>
                          </div>
                        </div>
                      ) : (
                        <div className="sale-payment-total">
                          <small>Total</small>
                          <strong>{money(s.total)}</strong>
                        </div>
                      )}

                      {s.paymentStatus !== 'pagada' && (
                        <button
                          type="button"
                          className="register-payment-btn"
                          onClick={() => openPaymentModal(s)}
                        >
                          <Plus size={15} />
                          Registrar pago
                        </button>
                      )}

                      {s.paymentStatus === 'pagada' &&
                        s.payments?.length === 1 && (
                          <span
                            className={
                              s.payments[0].paymentMethod === 'Mercado Pago'
                                ? 'mp'
                                : 'cash'
                            }
                          >
                            {s.payments[0].paymentMethod ===
                            'Mercado Pago' ? (
                              <WalletCards size={14} />
                            ) : (
                              <Banknote size={14} />
                            )}

                            {s.payments[0].paymentMethod}
                          </span>
                        )}

                        {(
                          s.paymentStatus === 'parcial' &&
                          s.payments?.length >= 1
                        ) || (
                          s.paymentStatus === 'pagada' &&
                          s.payments?.length > 1
                        ) ? (
                          <button
                            type="button"
                            className="view-payments-btn"
                            onClick={() => setPaymentHistoryModal(s)}
                          >
                            Ver pagos ({s.payments.length})
                          </button>
                        ) : null}

                      {isOwner && (
                        <small className="sale-profit">
                          Ganancia:{' '}
                          {s.profit !== null &&
                          s.profit !== undefined
                            ? money(s.profit)
                            : 'Pendiente'}
                        </small>
                      )}
                    </div>

                    <div className="sale-actions">
                      <button
                        onClick={() => openSaleEdit(s)}
                        title="Editar"
                      >
                        <Pencil size={15} />
                      </button>

                      <button
                        onClick={() => removeSale(s)}
                        title="Eliminar"
                      >
                        <Trash2 size={15} />
                      </button>
                    </div>
                  </article>
                ))}
              </div>
            )}

            <section className="history-panel">
              <div>
                <span className="eyebrow">
                  <History size={14} />
                  HISTORIAL
                </span>

                <h2>
                  Consultar ventas de otro día
                </h2>

                <p>
                  Elegí una fecha para revisar lo registrado.
                </p>
              </div>

              <div className="history-form">
                <input
                  type="date"
                  value={historyDate}
                  onChange={e =>
                    setHistoryDate(e.target.value)
                  }
                />

                <button
                  onClick={() => {
                    setSaleDate(historyDate);
                    setSalesView('day');
                    setSalesSearch('');
                  }}
                >
                  Ver día
                </button>
              </div>
            </section>
          </>
        )}
      </main>

      <footer>
        Distribuidora Bartol
        <span>•</span>
        Gestión simple para el día a día
      </footer>

      {logoutConfirm && (
        <div className="overlay">
          <div className="confirm">
            <button
              type="button"
              className="close"
              onClick={() => setLogoutConfirm(false)}
              title="Cerrar"
            >
              <X />
            </button>

            <h2>¿Cerrar sesión?</h2>

            <p>
              ¿Estás seguro de que querés cerrar la sesión?
            </p>

            <div className="modal-actions">
              <button
                className="cancel"
                onClick={() => setLogoutConfirm(false)}
              >
                Cancelar
              </button>

              <button
                className="delete"
                onClick={async () => {
                  setLogoutConfirm(false);
                  await logout();
                }}
              >
                Cerrar sesión
              </button>
            </div>
          </div>
        </div>
      )}

      {loginOpen && (
        <div className="overlay">
          <form
            className="modal"
            onSubmit={login}
          >
            <button
              type="button"
              className="close"
              onClick={() => {
                setLoginOpen(false);
                setLoginError('');
              }}
            >
              <X />
            </button>

            <div className="modal-icon">
              <Package />
            </div>

            <h2>Iniciar sesión</h2>

            <p>
              Acceso para personal de Distribuidora Bartol.
            </p>

            <label>
              Usuario

              <input
                autoFocus
                value={loginForm.username}
                onChange={e =>
                  setLoginForm({
                    ...loginForm,
                    username: e.target.value
                  })
                }
                placeholder="Usuario"
                autoComplete="username"
              />
            </label>

            <label>
              Contraseña

              <input
                type="password"
                value={loginForm.password}
                onChange={e =>
                  setLoginForm({
                    ...loginForm,
                    password: e.target.value
                  })
                }
                placeholder="Contraseña"
                autoComplete="current-password"
              />
            </label>

            {loginError && (
              <div className="login-error">
                <AlertCircle size={16} />
                {loginError}
              </div>
            )}

            <div className="modal-actions">
              <button
                type="button"
                className="cancel"
                onClick={() => {
                  setLoginOpen(false);
                  setLoginError('');
                }}
              >
                Cancelar
              </button>

              <button
                className="save"
                disabled={loginLoading}
              >
                {loginLoading
                  ? 'Ingresando...'
                  : 'Ingresar'}
              </button>
            </div>
          </form>
        </div>
      )}

      {restoreConfirm && (
        <div className="overlay">
          <div className="confirm">
            <button
              type="button"
              className="close"
              onClick={() => setRestoreConfirm(null)}
              title="Cerrar"
            >
              <X />
            </button>

            <div className="modal-icon">
              <Package />
            </div>

            <h2>¿Restaurar producto?</h2>

            <p>
              ¿Estás seguro de que querés restaurar{' '}
              <b>{restoreConfirm.name}</b>?
              El producto volverá a aparecer en su categoría original.
            </p>

            <div className="modal-actions">
              <button
                className="cancel"
                onClick={() => setRestoreConfirm(null)}
              >
                Cancelar
              </button>

              <button
                className="save"
                onClick={async () => {
                  const product = restoreConfirm;

                  setRestoreConfirm(null);

                  await restoreProduct(product);
                }}
              >
                Restaurar
              </button>
            </div>
          </div>
        </div>
      )}

      {stockModal && (
        <div className="overlay">
          <form
            className="modal"
            onSubmit={saveStock}
          >
            <button
              type="button"
              className="close"
              onClick={() => {
                setStockModal(null);
                setStockValue('');
              }}
            >
              <X />
            </button>

            <div className="modal-icon">
              <Package />
            </div>

            <h2>Editar stock</h2>

            <p>
              Producto: <b>{stockModal.name}</b>
            </p>

            <label>
              Stock disponible

              <input
                autoFocus
                type="number"
                min="0"
                step="1"
                value={stockValue}
                onChange={e => setStockValue(e.target.value)}
                onWheel={e => e.currentTarget.blur()}
              />
            </label>

            <div className="modal-actions">
              <button
                type="button"
                className="cancel"
                onClick={() => {
                  setStockModal(null);
                  setStockValue('');
                }}
              >
                Cancelar
              </button>

              <button className="save">
                Guardar stock
              </button>
            </div>
          </form>
        </div>
      )}

      {saleModal && (
        <div className="overlay">
          <form
            className="modal sale-modal"
            onSubmit={saveSale}
          >
            <button
              type="button"
              className="close"
              onClick={() => setSaleModal(null)}
            >
              <X />
            </button>

            <div className="modal-icon sale-icon">
              <ReceiptText />
            </div>

            <h2>
              {saleModal === 'add'
                ? 'Registrar venta'
                : 'Editar venta'}
            </h2>

            <p>
              Agregá todos los productos incluidos en la venta.
              Podés modificar el precio cobrado según el descuento
              acordado con el cliente.
            </p>

            {saleModal === 'add' && salesView === 'month' && (
              <label>
                Fecha de la venta

                <input
                  type="date"
                  value={saleRegisterDate}
                  min={`${salesMonth}-01`}
                  max={(() => {
                    const [year, month] = salesMonth
                      .split('-')
                      .map(Number);

                    const lastDay = new Date(
                      year,
                      month,
                      0
                    ).getDate();

                    return `${salesMonth}-${String(lastDay).padStart(
                      2,
                      '0'
                    )}`;
                  })()}
                  onChange={e =>
                    setSaleRegisterDate(e.target.value)
                  }
                  required
                />
              </label>
            )}

            <div className="sale-items-list">
              {saleItems.map((item, index) => {
                const search =
                  saleItemSearches[index] || '';

                const matches = products
                  .filter(p => p.active)
                  .filter(p =>
                    p.name
                      .toLowerCase()
                      .includes(search.toLowerCase())
                  )
                  .slice(0, 8);

                const selectedProduct = products.find(
                  p =>
                    String(p.id) ===
                    String(item.productId)
                );

                return (
                  <div
                    className="sale-item-card"
                    key={index}
                  >
                    <div className="sale-item-header">
                      <strong>
                        Producto {index + 1}
                      </strong>

                      {saleItems.length > 1 && (
                        <button
                          type="button"
                          className="sale-item-remove"
                          title="Quitar producto"
                          onClick={() =>
                            removeSaleItem(index)
                          }
                        >
                          <Trash2 size={16} />
                        </button>
                      )}
                    </div>

                    {!selectedProduct ? (
                      <label>
                        Producto

                        <div className="product-picker">
                          <div className="product-search">
                            <Search size={16} />

                            <input
                              value={search}
                              onChange={e =>
                                setSaleItemSearches(
                                  searches => ({
                                    ...searches,
                                    [index]: e.target.value
                                  })
                                )
                              }
                              placeholder="Buscar producto..."
                              autoComplete="off"
                            />
                          </div>

                          {search.trim() && (
                            <div className="product-options">
                              {matches.length === 0 ? (
                                <div className="product-no-results">
                                  No se encontró ningún producto.
                                </div>
                              ) : (
                                matches.map(p => (
                                  <button
                                    type="button"
                                    key={p.id}
                                    onClick={() => {
                                      updateSaleItem(
                                        index,
                                        'productId',
                                        p.id
                                      );

                                      updateSaleItem(
                                        index,
                                        'unitPrice',
                                        p.price
                                      );

                                      setSaleItemSearches(
                                        searches => ({
                                          ...searches,
                                          [index]: ''
                                        })
                                      );
                                    }}
                                  >
                                    <span>{p.name}</span>

                                    <small>
                                      {p.price === 0
                                        ? 'Precio pendiente'
                                        : money(p.price)}
                                    </small>
                                  </button>
                                ))
                              )}
                            </div>
                          )}
                        </div>
                      </label>
                    ) : (
                      <div className="selected-sale-product">
                        <div className="selected-sale-product-info">
                          <span>{selectedProduct.name}</span>

                          <small>
                            Stock: {selectedProduct.stock}
                          </small>
                        </div>

                        <button
                          type="button"
                          className="change-sale-product"
                          onClick={() => {
                            updateSaleItem(
                              index,
                              'productId',
                              ''
                            );

                            updateSaleItem(
                              index,
                              'unitPrice',
                              ''
                            );

                            setSaleItemSearches(
                              searches => ({
                                ...searches,
                                [index]: ''
                              })
                            );
                          }}
                        >
                          Cambiar producto
                        </button>
                      </div>
                    )}

                    <div className="two-fields">
                      <label>
                        Cantidad

                        <input
                          type="number"
                          min="1"
                          step="1"
                          value={item.quantity}
                          onChange={e =>
                            updateSaleItem(
                              index,
                              'quantity',
                              e.target.value
                            )
                          }
                          onWheel={e => e.currentTarget.blur()}
                        />
                      </label>

                      <label>
                        Precio cobrado c/u

                        <input
                          type="number"
                          min="0"
                          step="1"
                          value={item.unitPrice}
                          onChange={e =>
                            updateSaleItem(
                              index,
                              'unitPrice',
                              e.target.value
                            )
                          }
                          onWheel={e => e.currentTarget.blur()}
                        />
                      </label>
                    </div>

                    <div className="sale-item-subtotal">
                      <span>Subtotal</span>

                      <strong>
                        {money(
                          Number(item.quantity || 0) *
                          Number(item.unitPrice || 0)
                        )}
                      </strong>
                    </div>
                  </div>
                );
              })}
            </div>

            <button
              type="button"
              className="add-sale-item"
              onClick={addSaleItem}
            >
              <Plus size={17} />
              Agregar otro producto
            </button>

            <label>
              Descripción / detalle

              <input
                value={saleForm.description}
                onChange={e =>
                  setSaleForm({
                    ...saleForm,
                    description: e.target.value
                  })
                }
                placeholder="Ej. Marcelo - reparto"
                required
              />
            </label>

            {saleModal === 'add' && (
              <label>
                Estado del pago

                <div className="payment-options">
                  <button
                    type="button"
                    className={
                      saleForm.paymentStatus === 'pagada'
                        ? 'selected'
                        : ''
                    }
                    onClick={() =>
                      setSaleForm({
                        ...saleForm,
                        paymentStatus: 'pagada',
                        amountPaid: ''
                      })
                    }
                  >
                    Pagada completamente
                  </button>

                  <button
                    type="button"
                    className={
                      saleForm.paymentStatus === 'parcial'
                        ? 'selected'
                        : ''
                    }
                    onClick={() =>
                      setSaleForm({
                        ...saleForm,
                        paymentStatus: 'parcial'
                      })
                    }
                  >
                    Pago parcial
                  </button>

                  <button
                    type="button"
                    className={
                      saleForm.paymentStatus === 'pendiente'
                        ? 'selected'
                        : ''
                    }
                    onClick={() =>
                      setSaleForm({
                        ...saleForm,
                        paymentStatus: 'pendiente',
                        amountPaid: ''
                      })
                    }
                  >
                    Pago pendiente
                  </button>
                </div>
              </label>
            )}

            {saleModal === 'add' &&
              saleForm.paymentStatus === 'parcial' && (
                <label>
                  Monto abonado
                  <input
                    type="number"
                    min="0"
                    step="1"
                    value={saleForm.amountPaid}
                    onChange={e =>
                      setSaleForm({
                        ...saleForm,
                        amountPaid: e.target.value
                      })
                    }
                    required
                    onWheel={e => e.currentTarget.blur()}
                  />
                </label>
              )}

            {(
              (
                saleModal === 'add' &&
                saleForm.paymentStatus !== 'pendiente'
              ) ||
              (
                saleModal !== 'add' &&
                saleModal?.payments?.length === 1
              )
            ) && (
              <label>
                Forma de pago

                <div className="payment-options">
                  <button
                    type="button"
                    className={
                      saleForm.paymentMethod === 'Efectivo'
                        ? 'selected'
                        : ''
                    }
                    onClick={() =>
                      setSaleForm({
                        ...saleForm,
                        paymentMethod: 'Efectivo'
                      })
                    }
                  >
                    <Banknote size={18} />
                    Efectivo
                  </button>

                  <button
                    type="button"
                    className={
                      saleForm.paymentMethod === 'Mercado Pago'
                        ? 'selected'
                        : ''
                    }
                    onClick={() =>
                      setSaleForm({
                        ...saleForm,
                        paymentMethod: 'Mercado Pago'
                      })
                    }
                  >
                    <WalletCards size={18} />
                    Mercado Pago
                  </button>
                </div>
              </label>
            )}

            <div className="sale-total-box">
              <span>Total cobrado</span>

              <strong>{money(saleTotal)}</strong>
            </div>

            <div className="modal-actions">
              <button
                type="button"
                className="cancel"
                onClick={() => setSaleModal(null)}
              >
                Cancelar
              </button>

              <button className="save">
                {saleModal === 'add'
                  ? 'Registrar venta'
                  : 'Guardar cambios'}
              </button>
            </div>
          </form>
        </div>
      )}

      {paymentModal && (
        <div className="overlay">
          <form
            className="modal payment-modal"
            onSubmit={savePayment}
          >
            <button
              type="button"
              className="close"
              onClick={() => setPaymentModal(null)}
            >
              <X />
            </button>

            <div className="modal-icon">
              <Banknote />
            </div>

            <h2>Registrar pago</h2>

            <p>{paymentModal.description}</p>

            <div className="payment-modal-summary">
              <div>
                <span>Total</span>
                <strong>
                  {money(paymentModal.total)}
                </strong>
              </div>

              <div>
                <span>Pagado</span>
                <strong>
                  {money(paymentModal.paid)}
                </strong>
              </div>

              <div>
                <span>Pendiente</span>
                <strong>
                  {money(paymentModal.pending)}
                </strong>
              </div>
            </div>

            <label>
              Fecha del pago

              <input
                type="date"
                required
                value={paymentForm.date}
                onChange={e =>
                  setPaymentForm({
                    ...paymentForm,
                    date: e.target.value
                  })
                }
              />
            </label>

            <label>
              Monto

              <input
                type="number"
                min="1"
                max={paymentModal.pending}
                step="1"
                required
                value={paymentForm.amount}
                onChange={e =>
                  setPaymentForm({
                    ...paymentForm,
                    amount: e.target.value
                  })
                }
                placeholder={`Máximo ${money(
                  paymentModal.pending
                )}`}
                onWheel={e => e.currentTarget.blur()}
              />
            </label>

            <label>
              Forma de pago

              <div className="payment-options">
                <button
                  type="button"
                  className={
                    paymentForm.paymentMethod === 'Efectivo'
                      ? 'selected'
                      : ''
                  }
                  onClick={() =>
                    setPaymentForm({
                      ...paymentForm,
                      paymentMethod: 'Efectivo'
                    })
                  }
                >
                  <Banknote size={18} />
                  Efectivo
                </button>

                <button
                  type="button"
                  className={
                    paymentForm.paymentMethod === 'Mercado Pago'
                      ? 'selected'
                      : ''
                  }
                  onClick={() =>
                    setPaymentForm({
                      ...paymentForm,
                      paymentMethod: 'Mercado Pago'
                    })
                  }
                >
                  <WalletCards size={18} />
                  Mercado Pago
                </button>
              </div>
            </label>

            {paymentModal.payments?.length > 0 && (
              <div className="payment-history">
                <strong>Historial de pagos</strong>

                {paymentModal.payments.map(payment => (
                  <div
                    className="payment-history-row"
                    key={payment.id}
                  >
                    <div>
                      <span>
                        {dateShort(payment.date)}
                      </span>

                      <small>
                        {payment.paymentMethod}
                      </small>
                    </div>

                    <strong>
                      {money(payment.amount)}
                    </strong>
                  </div>
                ))}
              </div>
            )}

            <div className="modal-actions">
              <button
                type="button"
                className="cancel"
                onClick={() => setPaymentModal(null)}
              >
                Cancelar
              </button>

              <button
                type="submit"
                className="save"
              >
                Registrar pago
              </button>
            </div>
          </form>
        </div>
      )}

      {paymentHistoryModal && (
        <div className="overlay">
          <div className="modal payment-history-modal">
            <button
              type="button"
              className="close"
              onClick={() => setPaymentHistoryModal(null)}
            >
              <X />
            </button>

            <div className="modal-icon">
              <WalletCards />
            </div>

            <h2>Historial de pagos</h2>

            <p>
              {paymentHistoryModal.description}
            </p>

            <div className="payment-modal-summary">
              <div>
                <span>Total</span>
                <strong>
                  {money(paymentHistoryModal.total)}
                </strong>
              </div>

              <div>
                <span>Pagado</span>
                <strong>
                  {money(paymentHistoryModal.paid)}
                </strong>
              </div>

              <div>
                <span>Pendiente</span>
                <strong>
                  {money(paymentHistoryModal.pending)}
                </strong>
              </div>
            </div>

            <div className="payment-history">
              <strong>Pagos registrados</strong>

              {paymentHistoryModal.payments.map(payment => (
                <div
                  className="payment-history-row"
                  key={payment.id}
                >
                  <div>
                    <span>
                      {dateShort(payment.date)}
                    </span>

                    <small>
                      {payment.paymentMethod}
                    </small>
                  </div>

                  <div className="payment-history-actions">
                    <strong>
                      {money(payment.amount)}
                    </strong>

                    <button
                      type="button"
                      title="Editar medio de pago"
                      onClick={() =>
                        setEditingPayment({
                          ...payment,
                          saleId: paymentHistoryModal.id
                        })
                      }
                    >
                      <Pencil size={15} />
                    </button>
                  </div>
                </div>
              ))}
            </div>

            <div className="modal-actions">
              <button
                type="button"
                className="save"
                onClick={() => setPaymentHistoryModal(null)}
              >
                Cerrar
              </button>
            </div>
          </div>
        </div>
      )}


      {editingPayment && (
        <div className="overlay">
          <div className="modal">
            <button
              type="button"
              className="close"
              onClick={() => setEditingPayment(null)}
            >
              <X />
            </button>

            <div className="modal-icon">
              <Pencil />
            </div>

            <h2>Editar medio de pago</h2>

            <p>
              Corregí el medio utilizado para este pago.
            </p>

            <label>
              Importe

              <input
                type="text"
                value={money(editingPayment.amount)}
                disabled
              />
            </label>

            <label>
              Forma de pago

              <div className="payment-options">
                <button
                  type="button"
                  className={
                    editingPayment.paymentMethod === 'Efectivo'
                      ? 'selected'
                      : ''
                  }
                  onClick={() =>
                    setEditingPayment({
                      ...editingPayment,
                      paymentMethod: 'Efectivo'
                    })
                  }
                >
                  <Banknote size={18} />
                  Efectivo
                </button>

                <button
                  type="button"
                  className={
                    editingPayment.paymentMethod === 'Mercado Pago'
                      ? 'selected'
                      : ''
                  }
                  onClick={() =>
                    setEditingPayment({
                      ...editingPayment,
                      paymentMethod: 'Mercado Pago'
                    })
                  }
                >
                  <WalletCards size={18} />
                  Mercado Pago
                </button>
              </div>
            </label>

            <div className="modal-actions">
              <button
                type="button"
                className="cancel"
                onClick={() => setEditingPayment(null)}
              >
                Cancelar
              </button>

              <button
                type="button"
                className="save"
                onClick={savePaymentMethod}
              >
                Guardar
              </button>
            </div>
          </div>
        </div>
      )}


      {modal && (
        <div className="overlay">
          <form
            className="modal"
            onSubmit={save}
          >
            <button
              type="button"
              className="close"
              onClick={() => setModal(null)}
            >
              <X />
            </button>

            <div className="modal-icon">
              {modal === 'add' ? (
                <Plus />
              ) : (
                <Pencil />
              )}
            </div>

            <h2>
              {modal === 'add'
                ? 'Nuevo producto'
                : 'Editar producto'}
            </h2>

            <p>
              {modal === 'add'
                ? 'Agregá un producto a tu lista.'
                : 'Modificá los datos del producto.'}
            </p>

            <label>
              Nombre

              <input
                autoFocus
                value={form.name}
                onChange={e =>
                  setForm({
                    ...form,
                    name: e.target.value
                  })
                }
                placeholder="Ej. Miel de eucalipto 1 KG"
              />
            </label>

            <label>
              Precio

              <input
                type="number"
                min="0"
                value={form.price}
                onChange={e =>
                  setForm({
                    ...form,
                    price: e.target.value
                  })
                }
                placeholder="0"
                onWheel={e => e.currentTarget.blur()}
              />
            </label>

            {isOwner && (
              <label>
                Precio de Costo

                <input
                  type="number"
                  min="0"
                  placeholder="Costo Pendiente"
                  value={form.cost}
                  onChange={e =>
                    setForm({
                      ...form,
                      cost: e.target.value
                    })
                  }
                  onWheel={e => e.currentTarget.blur()}
                />
              </label>
            )}

            <label>
              Categoría

              <div className="select">
                <select
                  value={form.category}
                  onChange={e => {
                    setForm({
                      ...form,
                      category: e.target.value
                    });

                    if (e.target.value !== '__new__') {
                      setNewCategory('');
                    }
                  }}
                >
                  {categories
                    .filter(
                      c =>
                        c !== 'Todos' &&
                        c !== 'Productos eliminados'
                    )
                    .map(c => (
                      <option key={c}>
                        {c}
                      </option>
                    ))}

                  <option value="__new__">
                    ＋ Nueva categoría
                  </option>
                </select>

                <ChevronDown size={17} />
              </div>
            </label>

            {form.category === '__new__' && (
              <label>
                Nueva categoría

                <input
                  autoFocus
                  value={newCategory}
                  onChange={e =>
                    setNewCategory(e.target.value)
                  }
                  placeholder="Ej. Almacén"
                />
              </label>
            )}

            <div className="modal-actions">
              <button
                type="button"
                className="cancel"
                onClick={() => setModal(null)}
              >
                Cancelar
              </button>

              <button className="save">
                Guardar producto
              </button>
            </div>
          </form>
        </div>
      )}

      {confirm && (
        <div className="overlay">
          <div className="confirm">
            <div className="danger">
              <Trash2 />
            </div>

            <h2>
              ¿Eliminar{' '}
              {confirm.type === 'sale'
                ? 'venta'
                : 'producto'}
              ?
            </h2>

            <p>
              {confirm.type === 'sale' ? (
                <>
                  Vas a eliminar una venta con{' '}
                  <b>
                    {confirm.items.length}{' '}
                    {confirm.items.length === 1
                      ? 'producto'
                      : 'productos'}
                  </b>{' '}
                  por un total de{' '}
                  <b>{money(confirm.total)}</b>.
                  El stock de los productos será restaurado.
                </>
              ) : (
                <>
                  Vas a eliminar{' '}
                  <b>{confirm.name}</b>. Esta acción no se
                  puede deshacer.
                </>
              )}
            </p>

            <div className="modal-actions">
              <button
                className="cancel"
                onClick={() => setConfirm(null)}
              >
                Cancelar
              </button>

              <button
                className="delete"
                onClick={removeConfirmed}
              >
                Eliminar
              </button>
            </div>
          </div>
        </div>
      )}

      {notification && (
        <div className={`toast ${notification.type}`}>
          <CheckCircle2 size={18} />
          {notification.message}
        </div>
      )}
    </div>
  );
}

createRoot(document.getElementById('root')).render(<App />);