package grocery;
import static grocery.Main_Panel.isValidEmailFormat;
import java.awt.event.*;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.*;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.Date;
import java.util.logging.*;
import javax.mail.*;
import javax.mail.internet.*;
import javax.swing.*;
import javax.swing.Timer;
import javax.swing.table.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Functions
{
    Grocery grocery;
    Connection conn;
    DefaultTableModel table_product_function, table_customer_function, table_sales_function, table_checkout_function, table_items_sold_function;
    Session newSession = null;
    MimeMessage mimeMessage = null;
    Main_Panel mainPanel;
    Login_Panel loginPanel;
    String DATABASE_NAME = "inventory";
    
    Functions(Grocery grocery)
    {
        this.grocery = grocery;
        
    }
    
    public void main_Panel(Main_Panel mainPanel, Login_Panel loginPanel)
    {
        this.mainPanel = mainPanel;
        this.loginPanel = loginPanel;
    }
    
    public void connect()
    {
        try
        {
            conn = DriverManager.getConnection("jdbc:mysql://localhost:3306/", "root", "");
        } catch (SQLException ex) {
            Logger.getLogger(Login_Panel.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    public void validate(String username, String password)
    {   
        try {
            PreparedStatement stmt = conn.prepareStatement("SELECT * FROM users WHERE USERNAME = ? AND PASSWORD = ?");
            stmt.setString(1, username);
            stmt.setString(2, password);
            ResultSet rs = stmt.executeQuery();
            if (rs.next()) {
                grocery.show_Main_Panel();
            } else {
                JOptionPane.showMessageDialog(null, "Invalid username or password", "Error", JOptionPane.ERROR_MESSAGE);
                loginPanel.text.setText("");
                loginPanel.pass.setText("");
            }
        } catch (SQLException ex) {
            Logger.getLogger(Login_Panel.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    public void create_database() throws SQLException
    {
        try (
            Statement statement = conn.createStatement()) {

            String checkDatabaseQuery = "SHOW DATABASES LIKE '" + DATABASE_NAME + "'";
            if (!statement.executeQuery(checkDatabaseQuery).next()) {
                String createDatabaseQuery = "CREATE DATABASE " + DATABASE_NAME;
                statement.executeUpdate(createDatabaseQuery);
                
                create_table();
            }
            else
            {
                statement.execute("USE " + DATABASE_NAME);
            }
        }
    }
    
    private void create_table() throws SQLException
    {
        try (
            Statement statement = conn.createStatement()) {
            
            statement.execute("USE " + DATABASE_NAME);
            
            String createUserTable = "CREATE TABLE IF NOT EXISTS users ("
                    + "USER_ID INT AUTO_INCREMENT PRIMARY KEY, "
                    + "USERNAME VARCHAR(255), "
                    + "PASSWORD VARCHAR(255) "
                    + ")";
            statement.executeUpdate(createUserTable);
            
            String createProductTable = "CREATE TABLE IF NOT EXISTS products ("
                    + "PRODUCT_ID INT AUTO_INCREMENT PRIMARY KEY, "
                    + "PRODUCT_DESCRIPTION VARCHAR(255), "
                    + "PRODUCT_AVAILABLE_QUANTITY VARCHAR(255), "
                    + "PRODUCT_UNIT VARCHAR(255), "
                    + "PRODUCT_PRICE DOUBLE "
                    + ")";
            statement.executeUpdate(createProductTable);
            
            String createCustomerTable = "CREATE TABLE IF NOT EXISTS customer ("
                    + "CUSTOMER_ID INT AUTO_INCREMENT PRIMARY KEY, "
                    + "CUSTOMER_NAME VARCHAR(255), "
                    + "CUSTOMER_EMAIL VARCHAR(255) "
                    + ")";
            statement.executeUpdate(createCustomerTable);
            
            String createSalesTable = "CREATE TABLE IF NOT EXISTS sales ("
                    + "DATE DATE, "
                    + "TIME TIME, "
                    + "SALES_ID INT AUTO_INCREMENT PRIMARY KEY, "
                    + "CUSTOMER_ID INT, "
                    + "CUSTOMER_NAME VARCHAR(255), "
                    + "TOTAL_SALES DOUBLE, "
                    + "FOREIGN KEY (CUSTOMER_ID) REFERENCES customer(CUSTOMER_ID) "
                    + ")";
            statement.executeUpdate(createSalesTable);
            
            String createItemsSoldTable = "CREATE TABLE IF NOT EXISTS items_sold ("
                    + "SALES_ID INT, "
                    + "PRODUCT_ID INT, "
                    + "SOLD_QUANTITY INT, "
                    + "FOREIGN KEY (SALES_ID) REFERENCES sales(SALES_ID), "
                    + "FOREIGN KEY (PRODUCT_ID) REFERENCES products(PRODUCT_ID) "
                    + ")";
            statement.executeUpdate(createItemsSoldTable);
            
            PreparedStatement preparedStatement = conn.prepareStatement("INSERT INTO users (USERNAME, PASSWORD) VALUES (?, ?)");
            
            preparedStatement.setString(1, "Admin");
            preparedStatement.setString(2, "12345");

            preparedStatement.executeUpdate();
            
        }
    }
    
    public void create_product(String productDescription, int productAvailableQuantity, String productUnit, double productPrice)
    {
        try {
            PreparedStatement stmt = conn.prepareStatement("INSERT INTO products (PRODUCT_DESCRIPTION, PRODUCT_AVAILABLE_QUANTITY, PRODUCT_UNIT, PRODUCT_PRICE) VALUES (?, ?, ?, ?)");
            stmt.setString(1, productDescription);
            stmt.setInt(2, productAvailableQuantity);
            stmt.setString(3, productUnit);
            stmt.setDouble(4, productPrice);
            stmt.executeUpdate();
        } catch (SQLException ex) {
            Logger.getLogger(Main_Panel.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    public void create_customer(String customerName, String customerEmail)
    {
        try {
            PreparedStatement stmt = conn.prepareStatement("INSERT INTO customer (CUSTOMER_NAME, CUSTOMER_EMAIL) VALUES (?, ?)");
            stmt.setString(1, customerName);
            stmt.setString(2, customerEmail);
            stmt.executeUpdate();
        } catch (SQLException ex) {
            Logger.getLogger(Main_Panel.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    public void create_sales(int customerID, String name, double total)
    {
        try {
            PreparedStatement preparedStatement = conn.prepareStatement("INSERT INTO sales (DATE, TIME, CUSTOMER_ID, CUSTOMER_NAME, TOTAL_SALES) VALUES ( ?, ?, ?, ?, ?)");
            Date currentDate = new Date();
            SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
            SimpleDateFormat timeFormat = new SimpleDateFormat("HH:mm:ss");
            String dateString = dateFormat.format(currentDate);
            String timeString = timeFormat.format(currentDate);
            
            preparedStatement.setString(1, dateString);
            preparedStatement.setString(2, timeString);
            preparedStatement.setInt(3, customerID);
            preparedStatement.setString(4, name);
            preparedStatement.setDouble(5,total);

            preparedStatement.executeUpdate();
        } catch (SQLException ex) {
            Logger.getLogger(Functions.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    public void create_items_sold(int[] productID, int[] soldQuantity)
    {
        read_sales();
        String [][] salesData = new String [table_sales_function.getRowCount()][table_sales_function.getColumnCount()];
        int latestSalesID = 0;
        for(int i = 0; i < table_sales_function.getRowCount(); i++)
        {
            for(int j = 0; j < table_sales_function.getColumnCount(); j++)
            {
                salesData[i][j] = String.valueOf(table_sales_function.getValueAt(i,j));
            }
        }
        
        latestSalesID = Integer.parseInt(salesData[table_sales_function.getRowCount() - 1][2]);
        
        try {
            PreparedStatement preparedStatement = conn.prepareStatement("INSERT INTO items_sold (SALES_ID, PRODUCT_ID, SOLD_QUANTITY) VALUES (?, ?, ?)");

            conn.setAutoCommit(false);

            for(int i = 0;i < productID.length;i++)
            {
                preparedStatement.setInt(1, latestSalesID);
                preparedStatement.setInt(2, productID[i]);
                preparedStatement.setInt(3, soldQuantity[i]);
                preparedStatement.addBatch();
            }

            int [] count = preparedStatement.executeBatch();

            conn.commit();
        } catch (SQLException ex) {
            Logger.getLogger(Functions.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    public void read_product()
    {
        try {
            try (Statement statement = conn.createStatement()) {
                String query = "SELECT * FROM products";
                try (ResultSet resultSet = statement.executeQuery(query)) {
                    java.sql.ResultSetMetaData metaData = resultSet.getMetaData();
                    int columnCount = metaData.getColumnCount();
                    
                    int rowCount = table_product_function.getRowCount();
                    for(int i = rowCount - 1; i >= 0; i--)
                    {
                        table_product_function.removeRow(i);
                    }

                    while (resultSet.next()) {
                        Object[] row = new Object[columnCount];
                        for (int i = 1; i <= columnCount; i++) {
                            row[i - 1] = resultSet.getObject(i);
                        }
                        table_product_function.addRow(row);
                    }
                }
            }
        } catch (SQLException e) {
            JOptionPane.showMessageDialog(null, "Error fetching data from the database.");
        }
    }
    
    public void read_customer()
    {
        try {
            try (Statement statement = conn.createStatement()) {
                String query = "SELECT * FROM customer";
                try (ResultSet resultSet = statement.executeQuery(query)) {
                    java.sql.ResultSetMetaData metaData = resultSet.getMetaData();
                    int columnCount = metaData.getColumnCount();
                    
                    int rowCount = table_customer_function.getRowCount();
                    for(int i = rowCount - 1; i >= 0; i--)
                    {
                        table_customer_function.removeRow(i);
                    }

                    while (resultSet.next()) {
                        Object[] row = new Object[columnCount];
                        for (int i = 1; i <= columnCount; i++) {
                            row[i - 1] = resultSet.getObject(i);
                        }
                        table_customer_function.addRow(row);
                    }
                }
            }
        } catch (SQLException e) {
            JOptionPane.showMessageDialog(null, "Error fetching data from the database.");
        }
    }
    
    public void read_sales()
    {
        try {
            try (Statement statement = conn.createStatement()) {
                String query = "SELECT * FROM sales";
                try (ResultSet resultSet = statement.executeQuery(query)) {
                    java.sql.ResultSetMetaData metaData = resultSet.getMetaData();
                    int columnCount = metaData.getColumnCount();
                    
                    int rowCount = table_sales_function.getRowCount();
                    for(int i = rowCount - 1; i >= 0; i--)
                    {
                        table_sales_function.removeRow(i);
                    }

                    while (resultSet.next()) {
                        Object[] row = new Object[columnCount];
                        for (int i = 1; i <= columnCount; i++) {
                            row[i - 1] = resultSet.getObject(i);
                        }
                        table_sales_function.addRow(row);
                    }
                }
            }
        } catch (SQLException e) {
            JOptionPane.showMessageDialog(null, "Error fetching data from the database.");
        }
    }
    
    public void read_items_sold(int salesID, JTable product_table)
    {
        try {
            try (Statement statement = conn.createStatement()) {
                String query = "SELECT * FROM items_sold WHERE SALES_ID = " + salesID;

                try (ResultSet resultSet = statement.executeQuery(query)) {

                    table_items_sold_function.setRowCount(0);

                    while (resultSet.next()) {
                        String databaseProductID = resultSet.getString(2);
                        for (int i = 0; i < product_table.getRowCount(); i++) {
                            String productID = String.valueOf(product_table.getValueAt(i, 0));
                            if (databaseProductID.equals(productID)) {
                                Object[] row = new Object[6];
                                row[0] = product_table.getValueAt(i, 0);
                                row[1] = product_table.getValueAt(i, 1);
                                row[2] = resultSet.getObject(3);
                                row[3] = product_table.getValueAt(i, 3);
                                row[4] = product_table.getValueAt(i, 4);
                                row[5] = (Double.parseDouble(String.valueOf(resultSet.getObject(3))) * Double.parseDouble(String.valueOf(product_table.getValueAt(i, 4))));
                                table_items_sold_function.addRow(row);
                            }
                        }
                    }
                }
            }
        } catch (SQLException e) {
            JOptionPane.showMessageDialog(null, "Error fetching data from the database: " + e.getMessage());
        }
    }
    
    public void update(int productID, String productDescription, int productAvailableQuantity, String productUnit, double productPrice)
    {
        try {
            PreparedStatement stmt = conn.prepareStatement("UPDATE products SET PRODUCT_DESCRIPTION = ?, PRODUCT_AVAILABLE_QUANTITY = ?, PRODUCT_UNIT = ?, PRODUCT_PRICE = ? WHERE PRODUCT_ID = ?");
            stmt.setString(1, productDescription);
            stmt.setInt(2, productAvailableQuantity);
            stmt.setString(3, productUnit);
            stmt.setDouble(4, productPrice);
            stmt.setInt(5, productID);
            stmt.executeUpdate();
        } catch (SQLException ex) {
            Logger.getLogger(Main_Panel.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    public void delete(int productID)
    {
        try {
            PreparedStatement stmt = conn.prepareStatement("DELETE FROM products WHERE PRODUCT_ID = ?");
            stmt.setInt(1, productID);
            stmt.executeUpdate();
        } catch (SQLException ex) {
            Logger.getLogger(Main_Panel.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    private void setUpServerProperties()
    {
        Properties properties = System.getProperties();
        properties.put("mail.smtp.port","587");
        properties.put("mail.smtp.auth","true");
        properties.put("mail.smtp.starttls.enable","true");
        newSession = Session.getDefaultInstance(properties,null);
    }

    private void sendEmail() throws MessagingException
    {
        String fromUser = "elbernberdera@gmail.com"; //Pakibutang sa inyo Gmail diri
        String fromUserPassword = "jvsl fbuv dwaf ltdz"; // and App Password gikan sa inyong Gmail
        String emailHost = "smtp.gmail.com";
        try (Transport transport = newSession.getTransport("smtp"))
        {
            transport.connect(emailHost, fromUser, fromUserPassword);
            transport.sendMessage(mimeMessage, mimeMessage.getAllRecipients());
        }
        JOptionPane.showOptionDialog(null, "Email sent successfully.", "",
                JOptionPane.DEFAULT_OPTION,
                JOptionPane.INFORMATION_MESSAGE,
                null,
                new Object[]{},
                null);
    }
     
    protected MimeMessage draftEmail(int Customer_ID,String name, String email, String[][] order) throws AddressException, MessagingException, IOException
    {
        read_sales();
        int orderNumber = table_sales_function.getRowCount() + 1;
        
        String emailSubject = "Order #" + orderNumber;
        String orderList = "";
        double total = 0;
        for (String[] order1 : order)
        {
            for (int j = 0; j < order[0].length; j++)
            {
                if(j == 0)
                {
                    orderList += "\n<tr>";
                }
                orderList += "<td>" + order1[j] + "</td>";
                if(j == 4)
                {
                    orderList += "\n</tr>";
                    total += (Integer.valueOf(order1[2]) * Double.valueOf(order1[j]));
                }
            }
        }
        String emailBody = 
                "<html>Good Day " + name + ",<br><br>"
                + """
                Here are the list of items you have bought<br><br>
                <table style=\"width:100%\">
                    <tr>
                        <th>Product ID</th>
                        <th>Product Description</th>
                        <th>Quantity</th>
                        <th>Product Unit</th>
                        <th>Price</th>
                    </tr>
                """
                + orderList
                + """
                    <tr>
                        <th></th><th></th><th></th>
                        <th>TOTAL AMOUNT</th>
                """
                +      "<th>" + total + "</th>"
                + """
                    </tr>
                </table>
                <br><br>Thank you for purchasing with us!
                </html>
                """;
        mimeMessage = new MimeMessage(newSession);
        
        mimeMessage.addRecipient(Message.RecipientType.TO, new InternetAddress(email));
        mimeMessage.setSubject(emailSubject);

        MimeBodyPart bodyPart = new MimeBodyPart();
        bodyPart.setContent(emailBody,"text/html");
        MimeMultipart multiPart = new MimeMultipart();
        multiPart.addBodyPart(bodyPart);
        mimeMessage.setContent(multiPart);
        
        create_sales(Customer_ID, name, total);
        
        JOptionPane firstOptionPane = new JOptionPane("This might take a while",
                JOptionPane.INFORMATION_MESSAGE, JOptionPane.DEFAULT_OPTION, null, new Object[]{}, null);
        
        JDialog dialog = firstOptionPane.createDialog("Processing...");
        dialog.setDefaultCloseOperation(javax.swing.JDialog.DO_NOTHING_ON_CLOSE);
        Timer timer;
        timer = new Timer(2000, new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                dialog.dispose();
            }
        });

        timer.setRepeats(false);
        timer.start();

        dialog.setVisible(true);
        return mimeMessage;
    }
    
    public void purchase(JTable product_table,JTable checkout_table,JTable customer_table)
    {
        try {
            String name = "";
            String email = "";
            String[][] order = new String[checkout_table.getRowCount()][checkout_table.getColumnCount()];

            while (true) {
                name = JOptionPane.showInputDialog("Please enter your name:");
                if(name == null)
                {
                    break;
                }
                else if(name.isEmpty())
                {
                    JOptionPane.showMessageDialog(null, "Please fill out the window", "Error", JOptionPane.ERROR_MESSAGE);
                }
                else
                {
                    while (true)
                    {
                        email = JOptionPane.showInputDialog("Please enter your email address:");
                        if (email == null)
                        {
                            break;
                        }
                        else if (!isValidEmailFormat(email))
                        {
                            JOptionPane.showMessageDialog(null, "Please enter a valid email", "Error", JOptionPane.ERROR_MESSAGE);
                        }
                        else
                        {
                            break;
                        }
                    }
                    break;
                }
            }
            
            String [] customerRow = new String [3];
            int customerID = 0;
            boolean found = false;
            read_customer();
            if (name != null && !name.isEmpty() && email != null && !email.isEmpty())
            {
                for(int i = 0;i < customer_table.getRowCount();i++)
                {
                    for(int j = 0;j < customer_table.getColumnCount();j++)
                    {
                        customerRow[j] = String.valueOf(customer_table.getValueAt(i, j));
                    }

                    if(customerRow[2].equals(email))
                    {
                        customerID = Integer.parseInt(customerRow[0]);
                        found = true;
                    }
                }

                if(!found)
                {
                    System.out.println("Email saved into the database");
                    create_customer(name, email);
                    customerID += 1;
                }

                int[] productID_checkout = new int[checkout_table.getRowCount()];
                int[] soldQuantity = new int[checkout_table.getRowCount()];
                for(int i = 0;i < checkout_table.getRowCount();i++)
                {
                    for(int j = 0;j < checkout_table.getColumnCount();j++)
                    {
                        order[i][j] = String.valueOf(checkout_table.getValueAt(i, j));
                        if(j == 0)
                        {
                            productID_checkout[i] = Integer.parseInt(String.valueOf(checkout_table.getValueAt(i, j)));
                        }
                        else if(j == 2)
                        {
                            soldQuantity[i] = Integer.parseInt(String.valueOf(checkout_table.getValueAt(i, j)));
                        }
                    }
                }

                mainPanel.back_button[1].setEnabled(false);
                mainPanel.purchase_button[0].setEnabled(false);
                mainPanel.purchase_button[1].setEnabled(false);
                mainPanel.purchase_button[2].setEnabled(false);
                
                setUpServerProperties();
                draftEmail(customerID, name, email, order);
                create_items_sold(productID_checkout, soldQuantity);
                sendEmail();
                
                mainPanel.back_button[1].setEnabled(true);
                mainPanel.purchase_button[0].setEnabled(true);
                mainPanel.purchase_button[1].setEnabled(true);
                mainPanel.purchase_button[2].setEnabled(true);
                
                Object[][] orderData = new Object[order.length][order[0].length];

                for (int i = 0; i < order.length; i++) {
                    for (int j = 0; j < order[0].length; j++) {
                        orderData[i][j] = order[i][j];
                    }
                }
                
                Arrays.sort(orderData, (row1, row2) -> Integer.compare(
                                                       Integer.parseInt(String.valueOf(row1[0])),
                                                       Integer.parseInt(String.valueOf(row2[0]))));

                
                for(int i = 0; i < product_table.getRowCount();i++)
                {
                    String productID = String.valueOf(product_table.getValueAt(i, 0));
                    int productQuantity = Integer.parseInt(String.valueOf(product_table.getValueAt(i, 2)));
                    for(int j = 0;j < order.length;j++)
                    {
                        if(productID.equals(orderData[j][0]))
                        {
                            update(Integer.parseInt(String.valueOf(orderData[j][0])),
                            String.valueOf(orderData[j][1]),
                        productQuantity,
                                  String.valueOf(orderData[j][3]),
                                 Double.parseDouble(String.valueOf(orderData[j][4])));
                        }
                    }
                }
                
                if(checkout_table.getRowCount() != 0)
                {
                    for(int i = checkout_table.getRowCount() - 1; i >= 0;i--)
                    {
                        ((DefaultTableModel) checkout_table.getModel()).removeRow(i);
                    }
                }
            }
        }
        catch (MessagingException | IOException ex)
        {
            Logger.getLogger(Main_Panel.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    public void report() throws FileNotFoundException, IOException, java.text.ParseException
    {
        List<String> uniqueDates = getUniqueDatesFromSalesTable();
        Collections.sort(uniqueDates, Collections.reverseOrder());
        Date selectedDate = null;
        JComboBox<String> dateDropdown = new JComboBox<>(uniqueDates.toArray(new String[0]));

        JPanel panel = new JPanel();
        panel.add(new JLabel("Select Date:"));
        
        panel.add(dateDropdown);

        int result = JOptionPane.showConfirmDialog(null, panel, "Date Filter", JOptionPane.OK_CANCEL_OPTION,
                JOptionPane.PLAIN_MESSAGE);

        if (result == JOptionPane.OK_OPTION) {
            Object selectedItem = dateDropdown.getSelectedItem();
            if (selectedItem instanceof String string) {
                SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
                selectedDate = dateFormat.parse(string);
            }
            
            try {
                String fileName = "reports/" + dateDropdown.getSelectedItem() + " Sales Report.xlsx";
                XSSFWorkbook workbook = new XSSFWorkbook();

                createSheetForSales(workbook, "Sales", "SELECT * FROM sales WHERE DATE = ?", selectedDate);
                createSheet(workbook, "Customers", "SELECT * FROM customer");
                createSheetForItemsSold(workbook, "Items Sold", "SELECT * FROM items_sold", selectedDate);
                createSheet(workbook, "Products", "SELECT * FROM products");

                try (FileOutputStream outputStream = new FileOutputStream(fileName)) {
                    workbook.write(outputStream);

                    JOptionPane.showOptionDialog(null, "Excel file successfully created.", "",
                    JOptionPane.DEFAULT_OPTION,
                    JOptionPane.INFORMATION_MESSAGE,
                    null,
                    new Object[]{},
                    null);
                }
            } catch (SQLException ex) {
                Logger.getLogger(Main_Panel.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }
    
    private void createSheet(XSSFWorkbook workbook, String sheetName, String query) throws SQLException {
        PreparedStatement stmt = conn.prepareStatement(query);
        ResultSet rs = stmt.executeQuery();
        XSSFSheet sheet = workbook.createSheet(sheetName);

        int rowNum = 0;
        ResultSetMetaData metaData = rs.getMetaData();
        int columnCount = metaData.getColumnCount();

        XSSFRow headerRow = sheet.createRow(rowNum++);
        for (int i = 1; i <= columnCount; i++) {
            headerRow.createCell(i - 1).setCellValue(metaData.getColumnName(i));
        }

        while (rs.next()) {
            XSSFRow row = sheet.createRow(rowNum++);
            for (int i = 1; i <= columnCount; i++) {
                row.createCell(i - 1).setCellValue(rs.getObject(i).toString());
            }
        }
    }
    
    private void createSheetForSales(XSSFWorkbook workbook, String sheetName, String query, Date selectedDate) throws SQLException {
        try (PreparedStatement stmt = conn.prepareStatement(query)) {
            java.sql.Date sqlDate = new java.sql.Date(selectedDate.getTime());
            stmt.setDate(1, sqlDate);

            try (ResultSet rs = stmt.executeQuery()) {
                XSSFSheet sheet = workbook.createSheet(sheetName);

                int rowNum = 0;
                ResultSetMetaData metaData = rs.getMetaData();
                int columnCount = metaData.getColumnCount();

                XSSFRow headerRow = sheet.createRow(rowNum++);
                for (int i = 1; i <= columnCount; i++) {
                    headerRow.createCell(i - 1).setCellValue(metaData.getColumnName(i));
                }
                
                while (rs.next()) {
                    XSSFRow row = sheet.createRow(rowNum++);
                    for (int i = 1; i <= columnCount; i++) {
                        row.createCell(i - 1).setCellValue(rs.getObject(i).toString());
                    }
                }
            }
        }
    }

    private void createSheetForItemsSold(XSSFWorkbook workbook, String sheetName, String query, Date selectedDate) throws SQLException {
        PreparedStatement stmt = conn.prepareStatement(query);
        java.sql.Date sqlDate = new java.sql.Date(selectedDate.getTime());
        ResultSet rs = stmt.executeQuery();
        XSSFSheet sheet = workbook.createSheet(sheetName);

        int rowNum = 0;
        ResultSetMetaData metaData = rs.getMetaData();
        int columnCount = metaData.getColumnCount();

        XSSFRow headerRow = sheet.createRow(rowNum++);
        for (int i = 1; i <= columnCount; i++) {
            headerRow.createCell(i - 1).setCellValue(metaData.getColumnName(i));
        }

        while (rs.next()) {
            int salesID = rs.getInt("SALES_ID");
            if (salesIDExistsInSales(salesID, sqlDate)) {
                XSSFRow row = sheet.createRow(rowNum++);
                for (int i = 1; i <= columnCount; i++) {
                    row.createCell(i - 1).setCellValue(rs.getObject(i).toString());
                }
            }
        }
    }
   
    private boolean salesIDExistsInSales(int salesID, java.sql.Date selectedDate) throws SQLException {
        try (PreparedStatement stmt = conn.prepareStatement("SELECT * FROM sales WHERE SALES_ID = ? AND DATE = ?")) {
            stmt.setInt(1, salesID);
            stmt.setDate(2, selectedDate);
            try (ResultSet rs = stmt.executeQuery()) {
                return rs.next();
            }
        }
    }
    
    private List<String> getUniqueDatesFromSalesTable() {
        List<String> uniqueDates = new ArrayList<>();

        try (Statement statement = conn.createStatement();
             ResultSet resultSet = statement.executeQuery("SELECT DISTINCT DATE FROM sales")) {

            while (resultSet.next()) {
                uniqueDates.add(resultSet.getString("DATE"));
            }

        } catch (SQLException e) {
            e.printStackTrace();
        }

        return uniqueDates;
    }
}
