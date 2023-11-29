using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.IO;
using System.Windows.Forms;
using WindowsFormsApp1;

namespace Auto_Report.userControls
{
    public partial class form_1 : UserControl
    {
        Class1 class1 = new Class1();
        private int serialNumber = 1;
        private int overlapcount = 1;
        private DateTime overlapfinal;


        public form_1()
        {
            InitializeComponent();
            InitializeDataGridView();
            textBoxtotaldays.Visible = false;
            label17.Visible = false;
        }
        private void InitializeDataGridView()
        {
            dataGridView1.ColumnCount = 8;
            dataGridView1.Columns[0].Name = "Sno";
            dataGridView1.Columns[1].Name = "hindrance";
            dataGridView1.Columns[2].Name = "doh";
            dataGridView1.Columns[3].Name = "period";
            dataGridView1.Columns[4].Name = "extreq";
            dataGridView1.Columns[5].Name = "overlap";
            dataGridView1.Columns[6].Name = "netext";
            dataGridView1.Columns[7].Name = "remark";

        }

        //Total days calculating 
        private void btnTotalDays2_Click(object sender, EventArgs e)
        {
            DateTime selectedDate = dateTimePickerSA.Value;
            DateTime nextDay = selectedDate.AddDays(1);
            DateTime from = nextDay;
            DateTime to = dateTimePickerTo2.Value;

            TimeSpan diff = to - from;
            int totaldays = (int)diff.TotalDays;

            string total2 = totaldays.ToString();
        }

        //Total days calculating 
        private void btnTotaldays_Click(object sender, EventArgs e)
        {
            DateTime from = dateTimePickerAgreement.Value;
            DateTime to = dateTimePickerSA.Value;

            TimeSpan diff = to - from;
            int totaldays = (int)diff.TotalDays;
            string total1 = totaldays.ToString();
        }

        //Adding values to DataGrid view
        private void button1_Click(object sender, EventArgs e)
        {
            DateTime hold1;
            string textbox4;
            //remove time from datetime picker
            dateTimePicker1.Format = DateTimePickerFormat.Short;
            dateTimePicker2.Format = DateTimePickerFormat.Short;

            DateTime value1 = dateTimePicker1.Value;
            DateTime value2 = dateTimePicker2.Value;

            string selectedDate1 = value1.ToShortDateString();
            string selectedDate2 = value2.ToShortDateString();

            string check = class1.hold.ToString();

            int overlap = 0;

            string text;

            if (checkBox1.Checked)
            {
                text = "Attributed to Contractor";
            }
            else
            {
                text = "Not attributed to Contractor";
            }

            //calculate total no. of days
            DateTime from = dateTimePicker1.Value;
            DateTime to = dateTimePicker2.Value;

            TimeSpan diff = to - from;
            int totaldays = (int)diff.TotalDays + 1;
            int overlapdays = 0;

            if (check == "01-01-0001 00:00:00")
            {

            }
            else
            {
                TimeSpan diff1 = class1.hold - value2;
                overlap = (int)diff1.Days;
                overlapdays = totaldays + overlap;
            }

            int netextension = totaldays - overlapdays;

            //Adding values to DataGridView
            dataGridView1.Rows.Add(serialNumber++, textBox2.Text, selectedDate1, selectedDate2, selectedDate1 + " to" + selectedDate2 + "  total days= " + totaldays + " days", overlapdays + " days", netextension + " days", text);
            overlapcount++;

            hold1 = DateTime.Parse(selectedDate2);
            class1.hold = hold1;

            textBox2.Clear();
            int totaldayscount = 0;

            if (serialNumber == 1)
            {
                class1.totaldayscount = netextension;
                totaldayscount = netextension;
            }
            else
                totaldayscount = class1.totaldayscount + netextension;

            class1.totaldayscount = totaldayscount;
            textbox4 = totaldayscount.ToString();
            textBoxtotaldays.Text = totaldayscount.ToString();

        }


        private void addRow(String Sno, String hindrance, String doh, String period, String extreq, String overlap, String netext, String remark)
        {
            // add new row when btn is clicked
            String[] row = { Sno, hindrance, doh, period, extreq, overlap, netext, netext };
            dataGridView1.Rows.Add(row);
        }


        private void btnDSR_Click(object sender, EventArgs e)
        {
            int rowindex = dataGridView1.CurrentCell.RowIndex;
            dataGridView1.Rows.RemoveAt(rowindex);
        }

        //Export 
        private void button3_Click(object sender, EventArgs e)
        {
            PrintToPDF();
        }
        private void PrintToPDF()
        {


            try
            {

                string filePath = "form_details.pdf";

                // Create a PDF document

                var savefiledialoge = new SaveFileDialog();
                savefiledialoge.FileName = filePath;
                savefiledialoge.DefaultExt = ".pdf";
                if (savefiledialoge.ShowDialog() == DialogResult.OK)
                {
                    using (FileStream stream = new FileStream(savefiledialoge.FileName, FileMode.Create))
                    {
                        Document doc = new Document();
                        PdfWriter.GetInstance(doc, stream);
                        doc.Open();



                        #region Prt I

                        iTextSharp.text.Font boldFont = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA_BOLD, 9);
                        iTextSharp.text.Font boldFontlarge = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA_BOLD, 12);
                        iTextSharp.text.Font normalFont = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA, 9);
                        iTextSharp.text.Font redFont = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA, 9, BaseColor.RED);
                        iTextSharp.text.Font redFontbold = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA_BOLD, 9, BaseColor.RED);
                        // Add form details to PDF
                        iTextSharp.text.Paragraph title = new iTextSharp.text.Paragraph("APPLICATION FOR EXTENSION OF TIME (PART- I)", boldFontlarge); // Title in bold
                        title.Alignment = Element.ALIGN_CENTER; // Center alignment
                        doc.Add(title);
                        //add onr day
                        //total days 1
                        DateTime from = dateTimePickerAgreement.Value;
                        DateTime to = dateTimePickerSA.Value;

                        TimeSpan diff = to - from;
                        int totaldays = (int)diff.TotalDays;
                        string total1 = totaldays.ToString();

                        //datetime to short

                        dateTimePicker1.Format = DateTimePickerFormat.Short;
                        dateTimePicker2.Format = DateTimePickerFormat.Short;
                        dateTimePickerAgreement.Format = DateTimePickerFormat.Short;

                        dateTimePickerSA.Format = DateTimePickerFormat.Short;

                        //datetime to short
                        DateTime selectedDate = dateTimePickerSA.Value;
                        DateTime nextDay = selectedDate.AddDays(1);

                        dateTimePicker6.Format = DateTimePickerFormat.Short;
                        dateTimePicker5.Format = DateTimePickerFormat.Short;
                        dateTimePickerTo2.Format = DateTimePickerFormat.Short;
                        dateTimePicker3.Format = DateTimePickerFormat.Short;
                        dateTimePicker4.Format = DateTimePickerFormat.Short;

                        DateTime value1 = dateTimePicker1.Value;
                        DateTime value2 = dateTimePicker2.Value;
                        DateTime value5 = dateTimePicker5.Value;
                        DateTime value6 = dateTimePicker6.Value;
                        DateTime valueAgreement = dateTimePickerAgreement.Value;
                        DateTime valueFrom = dateTimePickerAgreement.Value;
                        DateTime valueTo = dateTimePickerSA.Value;
                        DateTime valueSA = dateTimePickerSA.Value;
                        DateTime valueFrom2 = nextDay;
                        DateTime valueTo2 = dateTimePickerTo2.Value;
                        DateTime valuedateproposed = dateTimePicker3.Value;
                        DateTime valuedatereceipt = dateTimePicker3.Value;

                        string selectedDate1 = value1.ToShortDateString();
                        string selectedDate2 = value2.ToShortDateString();
                        string selectedDate3 = valueAgreement.ToShortDateString();
                        string selectedDate4 = valueFrom.ToShortDateString();
                        string selectedDate5 = valueTo.ToShortDateString();
                        string selectedDate6 = valueSA.ToShortDateString();
                        string selectedDate7 = value6.ToShortDateString();
                        string selectedDate8 = value5.ToShortDateString();
                        string selectedDate9 = valueFrom2.ToShortDateString();
                        string selectedDate10 = valueTo2.ToShortDateString();
                        string selectedDate11 = valuedateproposed.ToShortDateString();
                        string selectedDate12 = valuedatereceipt.ToShortDateString();

                        iTextSharp.text.Paragraph linebreak = new iTextSharp.text.Paragraph("\n");
                        //----------------------------------------------------------------------------------------------------------------------------------------------------------------
                        doc.Add(linebreak);
                        doc.Add(linebreak);
                        doc.Add(linebreak);
                        doc.Add(linebreak);

                        PdfPTable tableform1 = new PdfPTable(3);
                        tableform1.WidthPercentage = 100;
                        float[] columnWidths = { 0.15f, 2f, 3.5f };
                        tableform1.SetWidths(columnWidths);
                        tableform1.SpacingBefore = 10f; // Adjust this value for space before each row
                        tableform1.SpacingAfter = 10f;


                        PdfPCell sno = new PdfPCell(new Phrase("1.", normalFont));
                        PdfPCell leftCellform1 = new PdfPCell(new Phrase("Name of Contractor: ", normalFont));
                        PdfPCell rightCellform1 = new PdfPCell(new Phrase(comboBoxCN.Text, boldFont));
                        
                        rightCellform1.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                        sno.Border = PdfPCell.NO_BORDER;
                        leftCellform1.Border = PdfPCell.NO_BORDER;
                        rightCellform1.Border = PdfPCell.NO_BORDER;
                        tableform1.AddCell(sno);
                        tableform1.AddCell(leftCellform1);
                        tableform1.AddCell(rightCellform1);
                        doc.Add(linebreak);

                        PdfPCell sno2 = new PdfPCell(new Phrase("2.", normalFont));
                        PdfPCell leftCellform2 = new PdfPCell(new Phrase("Name of the work as given in the Agreement:", normalFont));
                        PdfPCell rightCellform2 = new PdfPCell(new Phrase(txtBoxNA.Text, boldFont));
                       
                        rightCellform2.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                        sno2.Border = PdfPCell.NO_BORDER;
                        leftCellform2.Border = PdfPCell.NO_BORDER;
                        rightCellform2.Border = PdfPCell.NO_BORDER;
                        tableform1.AddCell(sno2);
                        tableform1.AddCell(leftCellform2);
                        tableform1.AddCell(rightCellform2);
                        doc.Add(linebreak);

                        PdfPCell sno3 = new PdfPCell(new Phrase("3.", normalFont));
                        PdfPCell leftCellform3 = new PdfPCell(new Phrase("Agreement No.:", normalFont));
                        PdfPCell rightCellform3 = new PdfPCell(new Phrase(txtBoxAN.Text, boldFont));
               
                        rightCellform2.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                        sno3.Border = PdfPCell.NO_BORDER;
                        leftCellform3.Border = PdfPCell.NO_BORDER;
                        rightCellform3.Border = PdfPCell.NO_BORDER;
                        tableform1.AddCell(sno3);
                        tableform1.AddCell(leftCellform3);
                        tableform1.AddCell(rightCellform3);
                        doc.Add(linebreak);

                        PdfPCell sno4 = new PdfPCell(new Phrase("4.", normalFont));
                        PdfPCell leftCellform4 = new PdfPCell(new Phrase("Estimated amount put to tender:", normalFont));
                        PdfPCell rightCellform4 = new PdfPCell(new Phrase(txtEAT.Text, boldFont));
            
                        rightCellform4.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                        sno4.Border = PdfPCell.NO_BORDER;
                        leftCellform4.Border = PdfPCell.NO_BORDER;
                        rightCellform4.Border = PdfPCell.NO_BORDER;
                        tableform1.AddCell(sno4);
                        tableform1.AddCell(leftCellform4);
                        tableform1.AddCell(rightCellform4);
                        doc.Add(linebreak);

                        PdfPCell sno5 = new PdfPCell(new Phrase("5.", normalFont));
                        PdfPCell leftCellform5 = new PdfPCell(new Phrase("Date of commencement of work as per Agreement:", normalFont));
                        PdfPCell rightCellform5 = new PdfPCell(new Phrase(selectedDate3, boldFont));

                        rightCellform5.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                        sno5.Border = PdfPCell.NO_BORDER;
                        leftCellform5.Border = PdfPCell.NO_BORDER;
                        rightCellform5.Border = PdfPCell.NO_BORDER;
                        tableform1.AddCell(sno5);
                        tableform1.AddCell(leftCellform5);
                        tableform1.AddCell(rightCellform5);
                        doc.Add(linebreak);

                        PdfPCell sno6 = new PdfPCell(new Phrase("6.", normalFont));
                        PdfPCell leftCellform6 = new PdfPCell( new iTextSharp.text.Phrase("Period allowed for completion of work as per Agreement:", normalFont));
                        PdfPCell rightCellform6 = new PdfPCell(new iTextSharp.text.Phrase(total1 + " days" + "( " + selectedDate4 + " to " + selectedDate5 + " )", boldFont));

                        rightCellform6.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                        sno6.Border = PdfPCell.NO_BORDER;
                        leftCellform6.Border = PdfPCell.NO_BORDER;
                        rightCellform6.Border = PdfPCell.NO_BORDER;
                        tableform1.AddCell(sno6);
                        tableform1.AddCell(leftCellform6);
                        tableform1.AddCell(rightCellform6);
                        doc.Add(linebreak);
    
                        PdfPCell sno7 = new PdfPCell(new Phrase("7.", normalFont));
                        PdfPCell leftCellform7 = new PdfPCell(new Phrase("State of completion stipulated in agreement:", normalFont));
                        PdfPCell rightCellform7 = new PdfPCell(new Phrase(selectedDate6, boldFont));

                        rightCellform7.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                        sno7.Border = PdfPCell.NO_BORDER;
                        leftCellform7.Border = PdfPCell.NO_BORDER;
                        rightCellform7.Border = PdfPCell.NO_BORDER;
                        tableform1.AddCell(sno7);
                        tableform1.AddCell(leftCellform7);
                        tableform1.AddCell(rightCellform7);
                        doc.Add(linebreak);

                        PdfPCell sno8 = new PdfPCell(new Phrase("8.", normalFont));
                        PdfPCell leftCellform8 = new PdfPCell(new Phrase("Period for which extension of time has been given previously:", normalFont));
                        PdfPCell rightCellform8 = new PdfPCell(new Phrase(txtBoxmonth.Text + " month's" + txtBoxdays.Text + " days", boldFont));

                        rightCellform8.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                        sno8.Border = PdfPCell.NO_BORDER;
                        leftCellform8.Border = PdfPCell.NO_BORDER;
                        rightCellform8.Border = PdfPCell.NO_BORDER;
                        tableform1.AddCell(sno8);
                        tableform1.AddCell(leftCellform8);
                        tableform1.AddCell(rightCellform8);

                        PdfPCell sno8a = new PdfPCell(new Phrase("", normalFont));
                        PdfPCell leftCellform8a = new PdfPCell(new Phrase("         a. 1st Extension vide EE,s :", normalFont));
                        PdfPCell rightCellform8a = new PdfPCell(new Phrase("", boldFont));

                        rightCellform8a.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                        sno8a.Border = PdfPCell.NO_BORDER;
                        leftCellform8a.Border = PdfPCell.NO_BORDER;
                        rightCellform8a.Border = PdfPCell.NO_BORDER;
                        tableform1.AddCell(sno8a);
                        tableform1.AddCell(leftCellform8a);
                        tableform1.AddCell(rightCellform8a);

                        PdfPCell sno8b = new PdfPCell(new Phrase("", normalFont));
                        PdfPCell leftCellform8b = new PdfPCell(new Phrase("        b. 2ndExtension vide EE,sNo.:", normalFont));
                        PdfPCell rightCellform8b = new PdfPCell(new Phrase("", boldFont));

                        rightCellform8b.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                        sno8b.Border = PdfPCell.NO_BORDER;
                        leftCellform8b.Border = PdfPCell.NO_BORDER;
                        rightCellform8b.Border = PdfPCell.NO_BORDER;
                        tableform1.AddCell(sno8b);
                        tableform1.AddCell(leftCellform8b);
                        tableform1.AddCell(rightCellform8b);

                        PdfPCell sno8c = new PdfPCell(new Phrase("", normalFont));
                        PdfPCell leftCellform8c = new PdfPCell(new Phrase("        c. 3rdExtension vide EE,sNo.:", normalFont));
                        PdfPCell rightCellform8c = new PdfPCell(new Phrase("", boldFont));

                        rightCellform8c.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                        sno8c.Border = PdfPCell.NO_BORDER;
                        leftCellform8c.Border = PdfPCell.NO_BORDER;
                        rightCellform8c.Border = PdfPCell.NO_BORDER;
                        tableform1.AddCell(sno8c);
                        tableform1.AddCell(leftCellform8c);
                        tableform1.AddCell(rightCellform8c);

                        PdfPCell sno8d = new PdfPCell(new Phrase("", normalFont));
                        PdfPCell leftCellform8d = new PdfPCell(new Phrase("        d. 4thExtension vide EE,sNo:", normalFont));
                        PdfPCell rightCellform8d = new PdfPCell(new Phrase("", boldFont));

                        rightCellform8d.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                        sno8d.Border = PdfPCell.NO_BORDER;
                        leftCellform8d.Border = PdfPCell.NO_BORDER;
                        rightCellform8d.Border = PdfPCell.NO_BORDER;
                        tableform1.AddCell(sno8d);
                        tableform1.AddCell(leftCellform8d);
                        tableform1.AddCell(rightCellform8d);

                        PdfPCell sno9 = new PdfPCell(new Phrase("9.", normalFont));
                        PdfPCell leftCellform9 = new PdfPCell(new Phrase("Reasons for which extension have been previously (Copies of the previous application should be attached) Period for which extension is applied for  :", normalFont));
                        PdfPCell rightCellform9 = new PdfPCell(new Phrase(selectedDate9 + " to " + selectedDate10 + " = " + total1 + " days.", boldFont));

                        rightCellform9.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                        sno9.Border = PdfPCell.NO_BORDER;
                        leftCellform9.Border = PdfPCell.NO_BORDER;
                        rightCellform9.Border = PdfPCell.NO_BORDER;
                        tableform1.AddCell(sno9);
                        tableform1.AddCell(leftCellform9);
                        tableform1.AddCell(rightCellform9);


                        doc.Add(tableform1);




                        //-----------------------------------------------------------------------------------------------------------------------------------
                        iTextSharp.text.Paragraph textField1 = new iTextSharp.text.Paragraph("1 Name of Contractor: ", normalFont);
                        iTextSharp.text.Paragraph textField1red = new iTextSharp.text.Paragraph(comboBoxCN.Text, redFontbold);

                        Phrase combinedText = new Phrase();
                        combinedText.Add(textField1);
                        combinedText.Add(" "); // Add space between the text boxes
                        combinedText.Add(textField1red);
                        doc.Add(combinedText);
                        
                        doc.Add(linebreak);

                        iTextSharp.text.Paragraph textField2 = new iTextSharp.text.Paragraph("2 Name of the work as given in the Agreement: -", normalFont);
                        iTextSharp.text.Paragraph textField2red = new iTextSharp.text.Paragraph(txtBoxNA.Text, redFontbold);

                        Phrase combinedText2 = new Phrase();
                        combinedText2.Add(textField2);
                        combinedText2.Add(" "); // Add space between the text boxes
                        combinedText2.Add(textField2red);
                        doc.Add(combinedText2);
                        doc.Add(linebreak);

                        iTextSharp.text.Paragraph textField3 = new iTextSharp.text.Paragraph("3 Agreement No.:", normalFont);
                        iTextSharp.text.Paragraph textField3red = new iTextSharp.text.Paragraph(txtBoxAN.Text, redFontbold);
                        Phrase combinedText3 = new Phrase();
                        combinedText3.Add(textField3);
                        combinedText3.Add(" "); // Add space between the text boxes
                        combinedText3.Add(textField3red);
                        doc.Add(combinedText3);
                        doc.Add(linebreak);

                        iTextSharp.text.Paragraph textField4 = new iTextSharp.text.Paragraph("4 Estimated amount put to tender:", normalFont);
                        iTextSharp.text.Paragraph textField4red = new iTextSharp.text.Paragraph(txtEAT.Text, redFontbold);
                        Phrase combinedText4 = new Phrase();
                        combinedText4.Add(textField4);
                        combinedText4.Add(" "); // Add space between the text boxes
                        combinedText4.Add(textField4red);
                        doc.Add(combinedText4);
                        doc.Add(linebreak);

                        iTextSharp.text.Paragraph textField5 = new iTextSharp.text.Paragraph("5 Date of commencement of work as per Agreement:", normalFont);
                        iTextSharp.text.Paragraph textField5red = new iTextSharp.text.Paragraph(selectedDate3, redFontbold);
                        Phrase combinedText5 = new Phrase();
                        combinedText5.Add(textField5);
                        combinedText5.Add(" "); // Add space between the text boxes
                        combinedText5.Add(textField5red);
                        doc.Add(combinedText5);
                        doc.Add(linebreak);


                        iTextSharp.text.Paragraph textField6 = new iTextSharp.text.Paragraph("6 Period allowed for completion of work as per Agreement:", normalFont);
                        iTextSharp.text.Paragraph textField6red = new iTextSharp.text.Paragraph(total1 + " days" + "( " + selectedDate4 + " to " + selectedDate5 + " )", redFontbold);
                        Phrase combinedText6 = new Phrase();
                        combinedText6.Add(textField6);
                        combinedText6.Add(" "); // Add space between the text boxes
                        combinedText6.Add(textField6red);
                        doc.Add(combinedText6);
                        doc.Add(linebreak);

                        iTextSharp.text.Paragraph textField7 = new iTextSharp.text.Paragraph("7 State of completion stipulated in agreement: ", normalFont);
                        iTextSharp.text.Paragraph textField7red = new iTextSharp.text.Paragraph(selectedDate6, boldFont);
                        Phrase combinedText7 = new Phrase();
                        combinedText7.Add(textField7);
                        combinedText7.Add(" "); // Add space between the text boxes
                        combinedText7.Add(textField7red);
                        doc.Add(combinedText7);
                        doc.Add(linebreak);

                        iTextSharp.text.Paragraph textField8 = new iTextSharp.text.Paragraph("8 Period for which extension of time has been given previously: " + txtBoxmonth.Text + " month's" + txtBoxdays.Text + " days", normalFont);
                        doc.Add(textField8);
                        doc.Add(linebreak);

                        iTextSharp.text.Paragraph textField9 = new iTextSharp.text.Paragraph("         a. 1st Extension vide EE,s " + txtBox8a.Text, normalFont);
                        doc.Add(textField9);
                        doc.Add(linebreak);

                        iTextSharp.text.Paragraph textField10 = new iTextSharp.text.Paragraph("        b. 2ndExtension vide EE,sNo.  " + txtBox8b.Text, normalFont);

                        doc.Add(textField10);
                        doc.Add(linebreak);

                        iTextSharp.text.Paragraph textField11 = new iTextSharp.text.Paragraph("        c. 3rdExtension vide EE,sNo.  " + txtBox8c.Text, normalFont);

                        doc.Add(textField11);
                        doc.Add(linebreak);

                        iTextSharp.text.Paragraph textField12 = new iTextSharp.text.Paragraph("        d. 4thExtension vide EE,sNo  " + txtBox8d.Text, normalFont);
                        doc.Add(textField12);
                        doc.Add(linebreak);


                        iTextSharp.text.Paragraph textField13 = new iTextSharp.text.Paragraph("9 Reasons for which extension have been previously (Copies of the previous application should be attached) Period for which extension is applied for  " + selectedDate9 + " to " + selectedDate10 + " = ", normalFont);
                        iTextSharp.text.Paragraph textField13bold = new iTextSharp.text.Paragraph(total1 + " days.", boldFontlarge);
                        Phrase combinedText13 = new Phrase();
                        combinedText13.Add(textField13);
                        combinedText13.Add(" "); // Add space between the text boxes
                        combinedText13.Add(textField13bold);
                        doc.Add(combinedText13);
                        doc.Add(linebreak);

                        iTextSharp.text.Paragraph textfield14 = new iTextSharp.text.Paragraph("10 Hindrance on account of which extension is applied for with dates on which hindrance occurred and the period for which these are likely to last.", normalFont);
                        doc.Add(textfield14);
                        doc.Add(linebreak);
                        doc.NewPage();
                        PdfPTable table = new PdfPTable(dataGridView1.Columns.Count);
                        table.WidthPercentage = 100;

                        // Add headers to PDF table
                        foreach (DataGridViewColumn col in dataGridView1.Columns)
                        {
                            PdfPCell headerCell = new PdfPCell(new Phrase(col.HeaderText, boldFont));
                            table.AddCell(headerCell);
                        }

                        // Add rows and cells to PDF table
                        foreach (DataGridViewRow row in dataGridView1.Rows)
                        {
                            foreach (DataGridViewCell cell in row.Cells)
                            {
                                PdfPCell cellToAdd = new PdfPCell(new Phrase(cell.Value?.ToString(), normalFont));
                                cellToAdd.MinimumHeight = 25f; // Set minimum height for cells

                                table.AddCell(cellToAdd);
                            }
                        }
                        doc.Add(table);

                        iTextSharp.text.Paragraph textField15 = new iTextSharp.text.Paragraph("Total  = " + textBoxtotaldays.Text + " days", boldFont);
                        textField15.Alignment = Element.ALIGN_RIGHT;
                        doc.Add(textField15);
                        doc.Add(linebreak);

                        iTextSharp.text.Paragraph textField16 = new iTextSharp.text.Paragraph("Total period for which extension is now Due to Water Logging applied for on account of above Hindrance:" + textBoxtotaldays.Text + " days", normalFont);
                        iTextSharp.text.Paragraph textField16bold = new iTextSharp.text.Paragraph(totaldays);
                        Phrase combinedText16 = new Phrase();
                        combinedText16.Add(textField16);
                        combinedText16.Add(" "); // Add space between the text boxes
                        combinedText16.Add(textField16bold);
                        doc.Add(combinedText16);
                        doc.Add(linebreak);

                        iTextSharp.text.Paragraph textField17 = new iTextSharp.text.Paragraph("11 Extension of time required for Extra work:" + textBox5.Text, normalFont);
                        doc.Add(textField17);
                        doc.Add(linebreak);

                        iTextSharp.text.Paragraph textField18 = new iTextSharp.text.Paragraph("12 Details of work and the amount involve:" + textBox6.Text, normalFont);
                        doc.Add(textField18);
                        doc.Add(linebreak);

                        iTextSharp.text.Paragraph textField19 = new iTextSharp.text.Paragraph("        a Total value of extra work: " + textBox7.Text, normalFont);
                        doc.Add(textField19);
                        doc.Add(linebreak);

                        iTextSharp.text.Paragraph textField20 = new iTextSharp.text.Paragraph("        b Proportionate period of extension of time based on estimate amount out to tender on account of extra work:" + textBox8.Text, normalFont);

                        doc.Add(textField20);
                        doc.Add(linebreak);

                        iTextSharp.text.Paragraph textField21 = new iTextSharp.text.Paragraph("13 Total extension of time required for the Months =" + textBoxtotaldays.Text + " days", normalFont);
                        doc.Add(textField21);
                        doc.Add(linebreak);

                        iTextSharp.text.Paragraph textField22 = new iTextSharp.text.Paragraph("14 Date of application Received by SO dated: -" + selectedDate7, normalFont);
                        doc.Add(textField22);
                        doc.Add(linebreak);


                        iTextSharp.text.Paragraph textField23 = new iTextSharp.text.Paragraph("15 Acknowledged by the SO vide his letter No " + selectedDate8, normalFont);
                        doc.Add(textField23);
                        doc.Add(linebreak);
                        doc.Add(linebreak);
                        doc.Add(linebreak);
                        doc.Add(linebreak);
                        doc.Add(linebreak);
                        doc.Add(linebreak);
                        doc.Add(linebreak);
                        doc.Add(linebreak);
                        doc.Add(linebreak);

                        PdfPTable tablelr = new PdfPTable(2);
                        table.WidthPercentage = 100;
                        PdfPCell leftCell = new PdfPCell(new Phrase("Submitted to the Sub divisional Officer", boldFontlarge));
                        PdfPCell rightCell = new PdfPCell(new Phrase("Signature of Contractor", boldFontlarge));

                       

                        leftCell.HorizontalAlignment = Element.ALIGN_LEFT;
                        rightCell.HorizontalAlignment = Element.ALIGN_RIGHT;

                        leftCell.Border = PdfPCell.NO_BORDER;
                        rightCell.Border = PdfPCell.NO_BORDER;

                        tablelr.AddCell(leftCell);
                        tablelr.AddCell(rightCell);

                        doc.Add(tablelr);

                        doc.Add(linebreak);
                        doc.Add(linebreak);
                        #endregion Part I


                        #region PartII

                        //--------------------------------------------------------------------------Second Form Part II-------------------------------------------------------------------------------------------------

                        doc.NewPage();
                        Paragraph title2 = new Paragraph("APPLICATION FOR EXTENSION OF TIME (PART- 2)", boldFontlarge); // Title in bold
                        title2.Alignment = Element.ALIGN_CENTER; // Center alignment
                        doc.Add(title2);
                        Paragraph subt = new Paragraph("(To be filled by the Sub-Divisional Officer)", boldFont);
                        subt.Alignment = Element.ALIGN_CENTER; // Center alignment
                        doc.Add(subt);

                        doc.Add(linebreak);

                        Paragraph txtfield = new Paragraph("1.Date of receipt of Application Form: " + textBox22.Text, normalFont);
                        doc.Add(txtfield);

                        doc.Add(linebreak);

                        dateTimePicker5.Format = DateTimePickerFormat.Short;
                        DateTime value7 = dateTimePicker5.Value;
                        string selectedDate17 = value7.ToShortDateString();

                        doc.Add(linebreak);

                        Paragraph txtfield2 = new Paragraph("2.Acknowledged issue by SDO vide his letter No : " + textBox3.Text + "Dated " + selectedDate17, normalFont);
                        doc.Add(txtfield2);

                        doc.Add(linebreak);

                        Paragraph txtfield3 = new Paragraph("3. Recommendation of S.D.O as to whether reasons given by the Contractor are correct and with extension, if any, is recommended by Him. If he does not recommend the extension reasons for rejection should be given. " + richTextBoxrecommendation.Text, normalFont);
                        doc.Add(txtfield3);

                        doc.Add(linebreak);
                        doc.Add(linebreak);
                        doc.Add(linebreak);
                        doc.Add(linebreak);

                        Paragraph txtfield4 = new Paragraph("To be filled by the Executive Engineer ", normalFont);
                        txtfield4.Alignment = Element.ALIGN_CENTER; // Center alignment
                        doc.Add(txtfield4);

                        doc.Add(linebreak);

                        Paragraph txtfield5 = new Paragraph("1. A.A of the project and the revised AA is required approved RD Dept. from: " + textBox10.Text, normalFont);
                        doc.Add(txtfield5);

                        doc.Add(linebreak);

                        Paragraph txtfield6 = new Paragraph("2. Amount of Administrative approval relating to the work for which tender is accepted.  " + textBox11.Text, normalFont);
                        doc.Add(txtfield6);

                        doc.Add(linebreak);


                        Paragraph txtfield7 = new Paragraph("3. Details of Technical sanction. Sanctioned vide Engineer-in-Chief Odisha/ SERWSS:  " + textBox12.Text, normalFont);
                        doc.Add(txtfield7);

                        doc.Add(linebreak);

                        Paragraph txtfield8 = new Paragraph("4. Accept tender amount / Amount of work put to tender  " + textBox13.Text, normalFont);
                        doc.Add(txtfield8);

                        doc.Add(linebreak);

                        Paragraph txtfield9 = new Paragraph("5.Split up Approval of any given by when and the No and Date of Tender:  " + textBox14.Text, normalFont);
                        doc.Add(txtfield9);

                        doc.Add(linebreak);

                        Paragraph txtfield10 = new Paragraph("6. Up to date Expenditure as against the amount for which tender is accepted:  " + textBox15.Text, normalFont);
                        doc.Add(txtfield10);

                        doc.Add(linebreak);

                        Paragraph txtfield11 = new Paragraph("7. Up to date Expenditure: " + textBox16.Text, normalFont);
                        doc.Add(txtfield11)
                            ;
                        doc.Add(linebreak);


                        Paragraph txtfield12 = new Paragraph("8. Proposed date of completion of the work : " + selectedDate11, normalFont);
                        doc.Add(txtfield12);

                        doc.Add(linebreak);

                        Paragraph txtfield13 = new Paragraph("9. Date of receipt in the Divisional Officer : " + selectedDate12, normalFont);
                        doc.Add(txtfield13);
                        doc.Add(linebreak);


                        Paragraph txtfield14 = new Paragraph("10. Executive Engineer’s regarding Hindrance mentioned by the Contractor :" + textBox19.Text, normalFont);
                        doc.Add(txtfield14);

                        doc.Add(linebreak);
                        doc.Add(linebreak);


                        // Add other form elements as needed               
                        doc.Add(table);
                        doc.Add(textField15);

                        doc.Add(linebreak);



                        Paragraph txtfield17 = new Paragraph("11. Executive engineer’s recommendations. The present progress of the work should be started and weather the work is likely to be completed by the date of Up to which extension has been applied for. If extension not recommended, what compensation is proposed to be levied under clause 2 of the agreement", normalFont);
                        doc.Add(txtfield17);

                        doc.Add(linebreak);

                        Paragraph txtfield18 = new Paragraph("    The present progress of the pipe water supply project to Village " + textBoxvillage.Text + " ," + textBoxblock.Text + " Block has been completed up to " + textboxpercent.Text + " %. The work is likely to be completed within the date applied for approval of EOT. As recommended by the Assistant Executive engineer RWS&S. Sub-division, Balasore as per Para 3.5.30.(i) of OPWD code Vol- I the EOT up to " + dateTimePickerTo2.Value + " may be approved without leave of penalty any financial benefit as the reasons stated by the Contractor is genuine.", normalFont);
                        doc.Add(txtfield18);

                        doc.Add(linebreak);
                        doc.Add(linebreak);
                        doc.Add(linebreak);
                        doc.Add(linebreak);
                        doc.Add(linebreak);

                        Paragraph txtfield19 = new Paragraph("  Executive Engineer Recommendations                                                     date of signature of Executive Engineer,", boldFont);
                        doc.Add(txtfield19);
                        Paragraph txtfield20 = new Paragraph("                                                                                                                                           RWS&S Division," + textBoxdistrict.Text, boldFont);
                        doc.Add(txtfield20);
                        doc.Add(linebreak);
                        doc.Add(linebreak);
                        doc.Add(linebreak);
                        doc.Add(linebreak);
                        doc.Add(linebreak);
                        Paragraph txtfield21 = new Paragraph("  Superintending Engineer Recommendations                                               date of signature of Superintending Engineer,", boldFont);
                        doc.Add(txtfield21);
                        Paragraph txtfield22 = new Paragraph("                                                                                                                                               RWS&SCircle," + textBoxdistrict.Text, boldFont);
                        doc.Add(txtfield22);
                        doc.Add(linebreak);
                        doc.Add(linebreak);
                        doc.Add(linebreak);
                        doc.Add(linebreak);
                        doc.Add(linebreak);
                        Paragraph txtfield23 = new Paragraph("  Chief-in- Engineer Recommendations                                                       date of signature of Chief-in-Engineer,", boldFont);
                        doc.Add(txtfield23);

                        Paragraph txtfield24 = new Paragraph("                                                                                                                                              RWS&S BBSR", boldFont);
                        doc.Add(txtfield24);
                        doc.Add(linebreak);
                        doc.Add(linebreak);

                        //-----------------------------------------------------------------No claim Certificates------------------------
                        doc.NewPage();



                        Chunk underlinedText = new Chunk("NO CLAIM CERTIFICATES", boldFontlarge);
                        underlinedText.SetUnderline(1.5f, -2); // Apply underline formatting

                        Paragraph heading = new Paragraph(underlinedText);
                        heading.Alignment = Element.ALIGN_CENTER;
                        doc.Add(heading);
                        doc.Add(linebreak);


                        Paragraph para = new Paragraph("    I shall not and do not claim compensation for the delay due to reason put forth by me for the work RPWS to village " + textBoxvillage.Text + " of " + textBoxblock.Text + " Block of " + textBoxdistrict.Text + " District. And stipulated time and date of completion for which extension time is prayed for now under clause-3 of G2   Agreement.", normalFont);
                        doc.Add(para);

                        doc.Add(linebreak);
                        doc.Add(linebreak);
                        doc.Add(linebreak);

                        Paragraph sign = new Paragraph("Signature of Contractor", boldFont);
                        sign.Alignment = Element.ALIGN_RIGHT;
                        doc.Add(sign);

                        doc.Add(linebreak);
                        doc.Add(linebreak);
                        doc.Add(linebreak);
                        doc.Add(linebreak);
                        doc.Add(linebreak);



                        Chunk underlinedText2 = new Chunk("NO CLAIM CERTIFICATES TOWARDS ESCALATION", boldFontlarge);
                        underlinedText2.SetUnderline(1.5f, -2); // Apply underline formatting
                        Paragraph heading2 = new Paragraph(underlinedText2);
                        heading2.Alignment = Element.ALIGN_CENTER;
                        doc.Add(heading2);
                        doc.Add(linebreak);


                        Paragraph para2 = new Paragraph("  I shall not and do not claim compensation for the delay due to reason put forth by me for the work RPWS to village " + textBoxvillage.Text + " of  " + textBoxblock.Text + " Block of " + textBoxdistrict.Text + " District. And stipulated time and date of completion for which extension time is prayed for now under clause-3 of G2 Agreement. ", normalFont);
                        doc.Add(para);

                        doc.Add(linebreak);
                        doc.Add(linebreak);
                        doc.Add(linebreak);

                        doc.Add(sign);



                        #endregion PartII


                        doc.Close();
                        stream.Close();

                        //PdfWriter.GetInstance(doc, new FileStream(filePath, FileMode.Create));
                    }
                }

                MessageBox.Show("PDF created successfully!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }

        }

       
    }
}

//PdfPTable tableform1 = new PdfPTable(3);
//tableform1.WidthPercentage = 100;
//PdfPCell sno = new PdfPCell(new Phrase(" ", boldFont));


//PdfPCell leftCellform = new PdfPCell(new Phrase(" ", normalFont));
//PdfPCell rightCellform1 = new PdfPCell(new Phrase(" ", boldFont));

////rightCell.HorizontalAlignment = Element.ALIGN_JUSTIFIED;

//sno.Border = PdfPCell.NO_BORDER;
//leftCellform.Border = PdfPCell.NO_BORDER;
//rightCellform.Border = PdfPCell.NO_BORDER;


//tableform.AddCell(sno);
//tableform.AddCell(leftCellform1);
//tableform.AddCell(rightCellform1);

//doc.Add(tableform1);



