package grocery;
import java.awt.*;
import java.awt.event.*;
import java.io.IOException;
import java.text.ParseException;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.*;
import javax.swing.table.*;

public class Main_Panel extends JPanel
{           
    JPanel[] main_panel = new JPanel[4];
    JPanel[] menu_panel = new JPanel[7];
    JPanel[] purchase_panel = new JPanel[4];
    JPanel[] bottom_panel = new JPanel[2];

    JSplitPane[] split_pane = new JSplitPane[2];
    JTabbedPane[] tabbed_pane = new JTabbedPane[2];
    JScrollPane[] scroll_pane = new JScrollPane[5];

    JButton[] main_button = new JButton[3];
    JButton[] menu_button = new JButton[5];
    JButton[] purchase_button = new JButton[3];
    JButton[] back_button = new JButton[2];

    JLabel[] main_label = new JLabel[4];
    JLabel[] menu_label = new JLabel[6];
    JLabel[] purchase_label = new JLabel[4];

    JTextField[] text = new JTextField[5];

    JTable product_table, customer_table, sales_table, checkout_table, items_sold_table;
    ListSelectionModel product_selection, customer_selection, sales_selection, checkout_selection, items_sold_selection;
    
    ImageIcon menu = new ImageIcon(Grocery.class.getResource("/resources/menu.png")), 
              purchase = new ImageIcon(Grocery.class.getResource("/resources/purchase.png")),
              logout = new ImageIcon(Grocery.class.getResource("/resources/logout.png")),
              add_menu = new ImageIcon(Grocery.class.getResource("/resources/add.png")), 
              add_purchase, 
              remove = new ImageIcon(Grocery.class.getResource("/resources/remove.png")), 
              checkout = new ImageIcon(Grocery.class.getResource("/resources/cart.png")), 
              update = new ImageIcon(Grocery.class.getResource("/resources/update.png")), 
              delete = new ImageIcon(Grocery.class.getResource("/resources/delete.png")), 
              records = new ImageIcon(Grocery.class.getResource("/resources/records.png")), 
              generate = new ImageIcon(Grocery.class.getResource("/resources/generate.png"));
    
    Image image_menu = menu.getImage(),
          image_purchase = purchase.getImage(),
          image_logout = logout.getImage(),
          image_add = add_menu.getImage(), 
          image_remove = remove.getImage(), 
          image_checkout = checkout.getImage(), 
          image_update = update.getImage(), 
          image_delete = delete.getImage(), 
          image_records = records.getImage(), 
          image_generate = generate.getImage();
    Image menu_resize = image_menu.getScaledInstance(90, 90, java.awt.Image.SCALE_SMOOTH),
          purchase_resize = image_purchase.getScaledInstance(90, 90, java.awt.Image.SCALE_SMOOTH),
          logout_resize = image_logout.getScaledInstance(90, 90, java.awt.Image.SCALE_SMOOTH),
          add_menu_resize = image_add.getScaledInstance(50, 50, java.awt.Image.SCALE_SMOOTH),
          update_resize = image_update.getScaledInstance(50, 50, java.awt.Image.SCALE_SMOOTH),
          delete_resize = image_delete.getScaledInstance(50, 50, java.awt.Image.SCALE_SMOOTH),
          records_resize = image_records.getScaledInstance(50, 50, java.awt.Image.SCALE_SMOOTH),
          generate_resize = image_generate.getScaledInstance(50, 50, java.awt.Image.SCALE_SMOOTH),
          add_purchase_resize = image_add.getScaledInstance(70, 70, java.awt.Image.SCALE_SMOOTH),
          remove_resize = image_remove.getScaledInstance(70, 70, java.awt.Image.SCALE_SMOOTH),
          checkout_resize = image_checkout.getScaledInstance(70, 70, java.awt.Image.SCALE_SMOOTH);
    
    Main_Panel(Grocery grocery, Functions function)
    {   
        menu = new ImageIcon(menu_resize); 
        purchase = new ImageIcon(purchase_resize);
        logout = new ImageIcon(logout_resize);
        add_menu = new ImageIcon(add_menu_resize);
        update = new ImageIcon(update_resize);
        delete = new ImageIcon(delete_resize);
        records = new ImageIcon(records_resize);
        generate = new ImageIcon(generate_resize);
        add_purchase = new ImageIcon(add_purchase_resize);
        remove = new ImageIcon(remove_resize);
        checkout = new ImageIcon(checkout_resize);
        
        setLayout(new BorderLayout(0, 0));
        for(int i = 0;i < 7;i++)
        {
            menu_panel[i] = new JPanel();
            if(i < 5)
            {
                text[i] = new JTextField();
                scroll_pane[i] = new JScrollPane();
                if (i < 4)
                {
                    main_panel[i] = new JPanel();
                    purchase_panel[i] = new JPanel();
                    if (i < 2)
                    {
                        split_pane[i] = new JSplitPane(JSplitPane.VERTICAL_SPLIT);
                        tabbed_pane[i] = new JTabbedPane();
                        bottom_panel[i] = new JPanel();
                    }
                }
            }
        }

        //Main Panel
        add(main_panel[0],BorderLayout.WEST);
        add(main_panel[1],BorderLayout.EAST);

        main_panel[0].setPreferredSize(new Dimension(200,100));
        main_panel[1].setPreferredSize(new Dimension(585,100));

        main_panel[0].setLayout(new BorderLayout(0, 0));
        main_panel[1].setLayout(new BorderLayout(0, 0));

        main_panel[0].add(main_panel[2]);
        main_panel[1].add(main_panel[3]);

        //Left
        main_panel[2].setBackground(Color.LIGHT_GRAY);
        main_panel[2].setLayout(null);
        main_panel[2].add(main_button[0] = new JButton(menu));
        main_panel[2].add(main_button[1] = new JButton(purchase));
        main_panel[2].add(main_button[2] = new JButton(logout));
        main_panel[2].add(main_label[0] = new JLabel("Grocery Menu"));
        main_panel[2].add(main_label[1] = new JLabel("Menu"));
        main_panel[2].add(main_label[2] = new JLabel("Purchase"));
        main_panel[2].add(main_label[3] = new JLabel("Log-out"));

        main_label[0].setFont(new java.awt.Font("Segoe UI", 1, 25));

        main_label[1].setFont(new java.awt.Font("Segoe UI", 1, 20));
        main_label[2].setFont(new java.awt.Font("Segoe UI", 1, 20));
        main_label[3].setFont(new java.awt.Font("Segoe UI", 1, 20));

        main_button[0].setBounds(50,80,100,100);
        main_button[1].setBounds(50,210,100,100);
        main_button[2].setBounds(50,340,100,100);
        
        main_label[0].setBounds(15,10,200,30);
        main_label[1].setBounds(70,50,100,30);
        main_label[2].setBounds(57,180,100,30);
        main_label[3].setBounds(61,310,100,30);

        main_button[0].addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e)
            {
                function.read_product();
                change_panel(menu_panel[0],split_pane());
            }
        });
        main_button[1].addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e)
            {
                function.read_product();

                tabbed_pane[0].add("Products",menu_panel[3]);
                tabbed_pane[0].add("Cart",purchase_panel[1]);

                clear();
                change_panel(purchase_panel[0], tabbed_pane[0]);
            }
        });
        main_button[2].addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e)
            {
                grocery.show_Login_Panel();
            }
        });

        //Right
        main_panel[3].setBackground(Color.DARK_GRAY);
        main_panel[3].setLayout(new BorderLayout(0, 0));

        //Menu Panel

        //Menu Panel Left
        menu_panel[0].setBackground(Color.LIGHT_GRAY);
        menu_panel[0].setLayout(null);

        menu_panel[0].add(menu_label[0] = new JLabel("Grocery Menu"));
        menu_panel[0].add(menu_label[1] = new JLabel("Add"));
        menu_panel[0].add(menu_label[2] = new JLabel("Update"));
        menu_panel[0].add(menu_label[3] = new JLabel("Delete"));
        menu_panel[0].add(menu_label[4] = new JLabel("Records"));
        menu_panel[0].add(menu_label[5] = new JLabel("<html>Generate<br>Excel</html>"));

        menu_panel[0].add(back_button[0] = new JButton("Back"));
        menu_panel[0].add(menu_button[0] = new JButton(add_menu));
        menu_panel[0].add(menu_button[1] = new JButton(update));
        menu_panel[0].add(menu_button[2] = new JButton(delete));
        menu_panel[0].add(menu_button[3] = new JButton(records));
        menu_panel[0].add(menu_button[4] = new JButton(generate));
        
        menu_label[0].setBounds(15,10,200,30);
        menu_label[1].setBounds(95,110,100,30);
        menu_label[2].setBounds(95,180,100,30);
        menu_label[3].setBounds(95,250,100,30);
        menu_label[4].setBounds(95,320,100,30);
        menu_label[5].setBounds(95,380,100,50);
        
        menu_label[0].setFont(new java.awt.Font("Segoe UI", 1, 25));
        menu_label[1].setFont(new java.awt.Font("Segoe UI", 1, 20));
        menu_label[2].setFont(new java.awt.Font("Segoe UI", 1, 20));
        menu_label[3].setFont(new java.awt.Font("Segoe UI", 1, 20));
        menu_label[4].setFont(new java.awt.Font("Segoe UI", 1, 20));
        menu_label[5].setFont(new java.awt.Font("Segoe UI", 1, 20));
        
        back_button[0].setBounds(45,50,110,30);
        menu_button[0].setBounds(25,100,60,60);
        menu_button[1].setBounds(25,170,60,60);
        menu_button[2].setBounds(25,240,60,60);
        menu_button[3].setBounds(25,310,60,60);
        menu_button[4].setBounds(25,380,60,60);
        
        back_button[0].addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                change_panel(main_panel[2],main_panel[3]);
                function.table_items_sold_function.setRowCount(0);
            }
        });
        menu_button[0].addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e)
            {               
                if(text[1].getText().matches("") &&
                   text[2].getText().matches("") &&
                   text[3].getText().matches("") &&
                   text[4].getText().matches(""))
                {
                    JOptionPane.showMessageDialog(null, "Please fill out the form", "Error", JOptionPane.ERROR_MESSAGE);
                }
                else if(productExists(text[1].getText()))
                {
                    JOptionPane.showMessageDialog(null, "Product already exist", "Error", JOptionPane.ERROR_MESSAGE);
                }
                else
                {
                    function.create_product(text[1].getText(),
                                Integer.parseInt(text[2].getText()),
                                          text[3].getText(),
                                         Double.parseDouble(text[4].getText()));
                    clear();
                }

                function.read_product();
                change_panel(null,split_pane());
                function.table_items_sold_function.setRowCount(0);
            }
        });
        menu_button[1].addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e)
            {   
                int row = 0;
                for(int i = 0; i < product_table.getRowCount(); i++)
                {
                    if(text[0].getText().matches(String.valueOf(product_table.getValueAt(i, 0))))
                    {
                        row = i;
                        break;
                    }
                }
                
                if(text[1].getText().matches("") &&
                   text[2].getText().matches("") &&
                   text[3].getText().matches("") &&
                   text[4].getText().matches(""))
                {
                    JOptionPane.showMessageDialog(null, "Please fill out the form", "Error", JOptionPane.ERROR_MESSAGE);
                }
                else if(text[0].getText().matches(String.valueOf(product_table.getValueAt(row,0))) != true)
                {
                    JOptionPane.showMessageDialog(null, "Product ID can't be changed", "Error", JOptionPane.ERROR_MESSAGE);
                }
                else if(text[1].getText().equals(product_table.getValueAt(row,1)) != true)
                {
                    JOptionPane.showMessageDialog(null, "Product Description can't be changed", "Error", JOptionPane.ERROR_MESSAGE);
                }
                else if(text[2].getText().equals(product_table.getValueAt(row, 2)) &&
                        text[3].getText().equals(product_table.getValueAt(row, 3)) &&
                        text[4].getText().equals(product_table.getValueAt(row, 4)))
                {
                    JOptionPane.showMessageDialog(null, "Please change something", "Error", JOptionPane.ERROR_MESSAGE);
                }
                else
                {
                    function.update(Integer.parseInt(text[0].getText()),
                             text[1].getText(),
                         Integer.parseInt(text[2].getText()),
                                   text[3].getText(),
                                  Double.parseDouble(text[4].getText()));
                    clear();
                }

                function.read_product();
                change_panel(null,split_pane());
                function.table_items_sold_function.setRowCount(0);
            }
        });
        menu_button[2].addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e)
            {
                if(text[0].getText().matches(""))
                {
                    JOptionPane.showMessageDialog(null, "Please fill out the form", "Error", JOptionPane.ERROR_MESSAGE);
                }
                else
                {
                    function.delete(Integer.parseInt(text[0].getText()));
                    clear();
                }

                function.read_product();
                change_panel(null,split_pane());
                function.table_items_sold_function.setRowCount(0);
            }
        });
        menu_button[3].addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e)
            {
                split_pane[1].setTopComponent(menu_panel[5]);
                split_pane[1].setBottomComponent(bottom_panel[1]);
                split_pane[1].getTopComponent();
                split_pane[1].getBottomComponent();
                split_pane[1].setDividerLocation(200);
                split_pane[1].setDividerSize(0);
                split_pane[1].setEnabled(false);

                menu_panel[6].add(split_pane[1]);

                function.read_customer();
                function.read_sales();

                tabbed_pane[1].add("Customers",menu_panel[4]);
                tabbed_pane[1].add("Sales",menu_panel[6]);

                clear();
                change_panel(menu_panel[0], tabbed_pane[1]);
                function.table_items_sold_function.setRowCount(0);
            }
        });
        menu_button[4].addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e)
            {
                try {
                    function.report();
                } catch (IOException | ParseException ex) {
                    Logger.getLogger(Main_Panel.class.getName()).log(Level.SEVERE, null, ex);
                }
            }
        });

        //Menu Panel Right
        menu_panel[1].setLayout(new BorderLayout(0, 0));
        menu_panel[2].setLayout(new BorderLayout(0, 0));
        menu_panel[6].setLayout(new BorderLayout(0, 0));
        bottom_panel[0].setLayout(null);

        bottom_panel[0].add(menu_label[0] = new JLabel("Product ID"));
        bottom_panel[0].add(menu_label[1] = new JLabel("Description"));
        bottom_panel[0].add(menu_label[2] = new JLabel("Available Quantity"));
        bottom_panel[0].add(menu_label[3] = new JLabel("Unit"));
        bottom_panel[0].add(menu_label[4] = new JLabel("Price"));

        for(int i = 0;i < 5;i++)
        {
            int x = 0, width = 0;
            bottom_panel[0].add(text[i]);
            switch(i)
            {
                case 0: x = 20; width = 65; break;
                case 1: x = 90; width = 150; break;
                case 2: x = 245; width = 110; break;
                case 3: x = 360; width = 50; break;
                case 4: x = 415; width = 150; break;
            }
            menu_label[i].setBounds(x,20,width,20);
            text[i].setBounds(x, 40, width, 20);
        }

        menu_panel[3].setLayout(new BorderLayout(0, 0));
        function.table_product_function = new DefaultTableModel() {
            {
                addColumn("Product ID");
                addColumn("Description");
                addColumn("Quantity");
                addColumn("Unit");
                addColumn("Price");
            }
        };
        product_table = new JTable(function.table_product_function);
        product_table.getColumnModel().getColumn(0).setPreferredWidth(50);
        product_table.getColumnModel().getColumn(1).setPreferredWidth(250);
        product_table.getColumnModel().getColumn(2).setPreferredWidth(50);
        product_table.getColumnModel().getColumn(3).setPreferredWidth(50);
        product_table.getColumnModel().getColumn(4).setPreferredWidth(50);
        product_table.getTableHeader().setReorderingAllowed(false);
        product_table.getTableHeader().setResizingAllowed(false);
        product_selection = product_table.getSelectionModel();
        product_selection.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        scroll_pane[0].setViewportView(product_table);
        menu_panel[3].add(scroll_pane[0]);

        menu_panel[4].setLayout(new BorderLayout(0,0));
        function.table_customer_function = new DefaultTableModel() {
            {
                addColumn("Customer ID");
                addColumn("Name");
                addColumn("Email");
            }
        };
        customer_table = new JTable(function.table_customer_function);
        customer_table.getColumnModel().getColumn(0).setPreferredWidth(50);
        customer_table.getColumnModel().getColumn(1).setPreferredWidth(150);
        customer_table.getColumnModel().getColumn(2).setPreferredWidth(200);
        customer_table.getTableHeader().setReorderingAllowed(false);
        customer_table.getTableHeader().setResizingAllowed(false);
        customer_selection = customer_table.getSelectionModel();
        customer_selection.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        scroll_pane[1].setViewportView(customer_table);
        menu_panel[4].add(scroll_pane[1], BorderLayout.CENTER);

        menu_panel[5].setLayout(new BorderLayout(0,0));
        function.table_sales_function = new DefaultTableModel() {
            {
                addColumn("Date");
                addColumn("Time");                
                addColumn("Sales ID");
                addColumn("Customer ID");
                addColumn("Name");
                addColumn("Total");
            }
        };
        sales_table = new JTable(function.table_sales_function);
        sales_table.getColumnModel().getColumn(0).setPreferredWidth(65);
        sales_table.getColumnModel().getColumn(1).setPreferredWidth(50);
        sales_table.getColumnModel().getColumn(2).setPreferredWidth(50);
        sales_table.getColumnModel().getColumn(3).setPreferredWidth(70);
        sales_table.getColumnModel().getColumn(4).setPreferredWidth(190);
        sales_table.getColumnModel().getColumn(5).setPreferredWidth(100);
        sales_table.getTableHeader().setReorderingAllowed(false);
        sales_table.getTableHeader().setResizingAllowed(false);
        sales_selection = sales_table.getSelectionModel();
        sales_selection.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        scroll_pane[2].setViewportView(sales_table);
        menu_panel[5].add(scroll_pane[2], BorderLayout.CENTER);
        
        bottom_panel[1].setLayout(new BorderLayout(0, 0));
        function.table_items_sold_function = new DefaultTableModel() {
            {
                addColumn("Product ID");
                addColumn("Description");
                addColumn("Quantity");
                addColumn("Unit");
                addColumn("Price");
                addColumn("Total Price");
            }
        };
        items_sold_table = new JTable(function.table_items_sold_function);
        items_sold_table.getColumnModel().getColumn(0).setPreferredWidth(50);
        items_sold_table.getColumnModel().getColumn(1).setPreferredWidth(250);
        items_sold_table.getColumnModel().getColumn(2).setPreferredWidth(45);
        items_sold_table.getColumnModel().getColumn(3).setPreferredWidth(45);
        items_sold_table.getColumnModel().getColumn(4).setPreferredWidth(50);
        items_sold_table.getColumnModel().getColumn(5).setPreferredWidth(60);
        items_sold_table.getTableHeader().setReorderingAllowed(false);
        items_sold_table.getTableHeader().setResizingAllowed(false);
        items_sold_selection = items_sold_table.getSelectionModel();
        items_sold_selection.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        scroll_pane[3].setViewportView(items_sold_table);
        bottom_panel[1].add(scroll_pane[3],BorderLayout.CENTER);
        
        product_table.setDefaultEditor(Object.class, null);
        customer_table.setDefaultEditor(Object.class, null);
        sales_table.setDefaultEditor(Object.class, null);
        items_sold_table.setDefaultEditor(Object.class, null);
        
        product_table.addMouseListener(new MouseAdapter()
        {
            @Override
            public void mouseClicked(MouseEvent e) {
                int row = product_table.rowAtPoint(e.getPoint());
                int col = product_table.columnAtPoint(e.getPoint());

                if (row >= 0 && col >= 0) {
                    for(int i = 0;i < 5;i++)
                    {
                        text[i].setText(String.valueOf(product_table.getValueAt(row, i)));
                    }
                }
            }
        });
        sales_table.addMouseListener(new MouseAdapter()
        {
            @Override
            public void mouseClicked(MouseEvent e) {
                int row = sales_table.rowAtPoint(e.getPoint());
                int col = sales_table.columnAtPoint(e.getPoint());

                if (row >= 0 && col >= 0) {    
                    function.read_items_sold(Integer.parseInt(String.valueOf(sales_table.getValueAt(row, 2))), product_table);
                }
            }
        });

        //Purchase Panel

        //Purchase Panel Left
        purchase_panel[0].setBackground(Color.LIGHT_GRAY);
        purchase_panel[0].setLayout(null);

        purchase_panel[0].add(back_button[1] = new JButton("Back"));
        purchase_panel[0].add(purchase_button[0] = new JButton(add_purchase));
        purchase_panel[0].add(purchase_button[1] = new JButton(remove));
        purchase_panel[0].add(purchase_button[2] = new JButton(checkout));
        purchase_panel[0].add(purchase_label[0] = new JLabel("Grocery Menu"));
        purchase_panel[0].add(purchase_label[1] = new JLabel("<html>Add<br>to cart</html>"));
        purchase_panel[0].add(purchase_label[2] = new JLabel("<html>Remove<br>from cart</html>"));
        purchase_panel[0].add(purchase_label[3] = new JLabel("Checkout"));
        
        purchase_label[0].setFont(new java.awt.Font("Segoe UI", 1, 25));
        purchase_label[1].setFont(new java.awt.Font("Segoe UI", 1, 20));
        purchase_label[2].setFont(new java.awt.Font("Segoe UI", 1, 20));
        purchase_label[3].setFont(new java.awt.Font("Segoe UI", 1, 20));
        
        purchase_label[0].setBounds(15,10,200,30);
        purchase_label[1].setBounds(100,130,100,50);
        purchase_label[2].setBounds(100,240,100,50);
        purchase_label[3].setBounds(100,350,100,50);

        back_button[1].setBounds(45,50,110,30);
        purchase_button[0].setBounds(15,120,80,80);
        purchase_button[1].setBounds(15,230,80,80);
        purchase_button[2].setBounds(15,340,80,80);

        back_button[1].addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e)
            {
                if(checkout_table.getRowCount() != 0)
                {
                    for(int i = checkout_table.getRowCount() - 1; i >= 0;i--)
                    {
                        ((DefaultTableModel) checkout_table.getModel()).removeRow(i);
                    }
                }
                
                change_panel(main_panel[2],main_panel[3]);
            }
        });
        purchase_button[0].addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e)
            {
                int selectedRow = product_table.getSelectedRow();
                if (selectedRow >= 0)
                {
                    String[] rowData = new String[5];
                    String quantity = null;

                    while (true)
                    {
                        String input = JOptionPane.showInputDialog("How Many?");
                        if (input != null)
                        {
                            try
                            {
                                int enteredQuantity = Integer.parseInt(input);

                                if (enteredQuantity <= Integer.parseInt(String.valueOf(product_table.getValueAt(selectedRow, 2))))
                                {
                                    quantity = input;
                                    break;
                                }
                                else
                                {
                                    JOptionPane.showMessageDialog(null, "Please enter a valid number", "Error", JOptionPane.ERROR_MESSAGE);
                                }
                            }
                            catch (NumberFormatException ex)
                            {
                                JOptionPane.showMessageDialog(null, "Please enter a valid number", "Error", JOptionPane.ERROR_MESSAGE);
                            }
                        }
                        else
                        {
                            break;
                        }
                    }
                    
                    if(quantity != null)
                    {
                        for (int i = 0; i < 5; i++)
                        {
                            rowData[i] = (i == 2) ? quantity : String.valueOf(product_table.getValueAt(selectedRow, i));
                        }

                        boolean found = false;
                        int checkoutQuantity;
                        int productQuantity;
                        
                        for (int row = 0; row < checkout_table.getRowCount(); row++)
                        {
                            int tableProductId = Integer.parseInt(String.valueOf(checkout_table.getValueAt(row, 0)));
                            if (tableProductId == Integer.parseInt(rowData[0]))
                            {
                                checkoutQuantity = Integer.parseInt(String.valueOf(checkout_table.getValueAt(row, 2)));
                                checkout_table.setValueAt(checkoutQuantity + Integer.parseInt(quantity), row, 2);
                                found = true;
                                break;
                            }
                        }
                        
                        productQuantity = Integer.parseInt(String.valueOf(product_table.getValueAt(selectedRow, 2)));
                        product_table.setValueAt(productQuantity - Integer.parseInt(quantity), selectedRow, 2);
                        if (!found)
                        {
                            ((DefaultTableModel) checkout_table.getModel()).addRow(rowData);
                        }
                    }
                }
                else
                {
                    JOptionPane.showMessageDialog(null, "Please select a row from the Products List", "Error", JOptionPane.ERROR_MESSAGE);
                }
            }
        });
        purchase_button[1].addActionListener(new ActionListener(){
            @Override
            public void actionPerformed(ActionEvent e)
            {
                int selectedRow = checkout_table.getSelectedRow();
                
                if (selectedRow >= 0)
                {    
                    String[] rowData = new String[5];
                    
                    for (int i = 0; i < 5; i++)
                    {
                        rowData[i] = String.valueOf(checkout_table.getValueAt(selectedRow, i));
                    }
                    
                    int tableProductId = Integer.parseInt(rowData[0]);
                    
                    for (int row = 0; row < product_table.getRowCount(); row++)
                    {
                        int currentProductId = Integer.parseInt(String.valueOf(product_table.getValueAt(row, 0)));
                        if (currentProductId == tableProductId)
                        {
                            int currentQuantity = Integer.parseInt(String.valueOf(product_table.getValueAt(row, 2)));
                            product_table.setValueAt(currentQuantity + Integer.parseInt(rowData[2]), row, 2);
                            ((DefaultTableModel) checkout_table.getModel()).removeRow(selectedRow);
                            break;
                        }
                    }
                }
                else
                {
                    JOptionPane.showMessageDialog(null, "Please select a row in from your checkout table", "Error", JOptionPane.ERROR_MESSAGE);
                }
            }
        });
        purchase_button[2].addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e)
            {
                int checkout = checkout_table.getRowCount();
                if(checkout == 0)
                {
                    JOptionPane.showMessageDialog(null, "Your cart is empty");
                }
                else
                {
                    function.purchase(product_table, checkout_table,customer_table);
                }
            }
        });
        
        //Purchase Panel Right
        purchase_panel[1].setBackground(Color.DARK_GRAY);
        purchase_panel[1].setLayout(new BorderLayout(0,0));
        
        function.table_checkout_function = new DefaultTableModel()
        {
            {
                addColumn("Product ID");
                addColumn("Description");
                addColumn("Quantity");
                addColumn("Unit");
                addColumn("Price");
            }
            
            @Override
            public boolean isCellEditable(int row, int column)
            {
                return column == 2;
            }
        };
        checkout_table = new JTable(function.table_checkout_function);
        checkout_table.getColumnModel().getColumn(0).setPreferredWidth(50);
        checkout_table.getColumnModel().getColumn(1).setPreferredWidth(250);
        checkout_table.getColumnModel().getColumn(2).setPreferredWidth(50);
        checkout_table.getColumnModel().getColumn(3).setPreferredWidth(50);
        checkout_table.getColumnModel().getColumn(4).setPreferredWidth(50);
        checkout_table.getTableHeader().setReorderingAllowed(false);
        checkout_table.getTableHeader().setResizingAllowed(false);
        checkout_selection = checkout_table.getSelectionModel();
        checkout_selection.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        checkout_table.setDefaultEditor(Object.class, null);
        scroll_pane[4].setViewportView(checkout_table);
        purchase_panel[1].add(scroll_pane[4], BorderLayout.CENTER);
        
        for(int i = 0; i < 5; i++)
        {
            menu_button[i].setBackground(new Color(0, 0, 0, 0));
            menu_button[i].setForeground(Color.WHITE);
            menu_button[i].setContentAreaFilled(false);
            menu_button[i].setOpaque(false);
            if(i < 3)
            {
                main_button[i].setBackground(new Color(0, 0, 0, 0));
                main_button[i].setForeground(Color.WHITE);
                main_button[i].setContentAreaFilled(false);
                main_button[i].setOpaque(false);
                
                purchase_button[i].setBackground(new Color(0, 0, 0, 0));
                purchase_button[i].setForeground(Color.WHITE);
                purchase_button[i].setContentAreaFilled(false);
                purchase_button[i].setOpaque(false);
            }
        }
    }

    protected static boolean isValidEmailFormat(String email)
    {
        String emailRegex = "^[a-zA-Z0-9_+&*-]+(?:\\.[a-zA-Z0-9_+&*-]+)*@(?:[a-zA-Z0-9-]+\\.)+[a-zA-Z]{2,7}$";
        return email.matches(emailRegex);
    }
    
    private boolean productExists(String productName)
    {
        for (int row = 0; row < product_table.getRowCount(); row++)
        {
            Object value = product_table.getValueAt(row, 1);
            if (value != null && value.toString().equalsIgnoreCase(productName))
            {
                return true;
            }
        }
        return false;
    }

    private JPanel split_pane()
    {
        split_pane[0].setTopComponent(menu_panel[3]);
        split_pane[0].setBottomComponent(bottom_panel[0]);
        split_pane[0].getTopComponent();
        split_pane[0].getBottomComponent();
        split_pane[0].setDividerLocation(360);
        split_pane[0].setDividerSize(0);
        split_pane[0].setEnabled(false);
        menu_panel[2].add(split_pane[0]);

        return menu_panel[2];
    }

    private void change_panel(JPanel panel_1, JPanel panel_2)
    {
        if (panel_1 != null)
        {
            main_panel[0].removeAll();
            main_panel[0].add(panel_1);
            main_panel[0].revalidate();
            main_panel[0].repaint();
        }
        main_panel[1].removeAll();
        main_panel[1].add(panel_2);
        main_panel[1].revalidate();
        main_panel[1].repaint();
    }

    private void change_panel(JPanel panel_1, JTabbedPane panel_2)
    {
        if (panel_1 != null)
        {
            main_panel[0].removeAll();
            main_panel[0].add(panel_1);
            main_panel[0].revalidate();
            main_panel[0].repaint();
        }
        main_panel[1].removeAll();
        main_panel[1].add(panel_2);
        main_panel[1].revalidate();
        main_panel[1].repaint();
    }

    private void clear()
    {
        for(int i = 0;i < 5;i++)
        {
            text[i].setText("");
        }
    }
}