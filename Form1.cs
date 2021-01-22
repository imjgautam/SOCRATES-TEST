using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections;
using iTextSharp.text;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.IO;
using System.Runtime.InteropServices;
using System.Drawing.Design;
using Microsoft.Office.Interop.Excel;
using word = Microsoft.Office.Interop.Word;   




using iTextSharp.text.pdf;
using Application = System.Windows.Forms.Application;

namespace Test_1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            btn_submit.Enabled = false;
            bt_amb_submit.Enabled = false;
            btn_ts_submit.Enabled = false;

            btn_refresh_2.Enabled = false;
            btn_refresh_3.Enabled = false;
            btn_refresh.Enabled = false;
            btn_result_print.Enabled = false;
            //  groupBox1.Enabled = false;
            // groupBox2.Enabled = false;
            // groupBox3.Enabled = false;
            //ArrayList al = new ArrayList();
        }

        private void TabPage1_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void TableLayoutPanel2_Paint(object sender, PaintEventArgs e)
        {

        }




        public string GetSelectedRadioButtonText(Panel grb)
        {
            return grb.Controls.OfType<RadioButton>().SingleOrDefault(rad => rad.Checked == true).Text;
        }


        private void Btn_submit_Click(object sender, EventArgs e)
        {
            // int value = Convert.ToInt32(this.lbl_r_total.Text);
            txt_1q.Text = GetSelectedRadioButtonText(panel1);
            txt_2q.Text = GetSelectedRadioButtonText(panel2);

            txt_3q.Text = GetSelectedRadioButtonText(panel3);

            txt_4q.Text = GetSelectedRadioButtonText(panel4);

            txt_5q.Text = GetSelectedRadioButtonText(panel5);


            txt_6q.Text = GetSelectedRadioButtonText(panel6);

            txt_7q.Text = GetSelectedRadioButtonText(panel7);



            lbl_r_total.Text = Convert.ToString(Convert.ToInt32(txt_1q.Text) + Convert.ToInt32(txt_2q.Text) +
               Convert.ToInt32(txt_3q.Text) + Convert.ToInt32(txt_4q.Text) + Convert.ToInt32(txt_5q.Text) + Convert.ToInt32(txt_6q.Text)
               + Convert.ToInt32(txt_7q.Text));


            MessageBox.Show("Submitted successfuly");
            btn_submit.Enabled = false;
            btn_refresh.Enabled = true;
        }
        private void Bt_amb_submit_Click(object sender, EventArgs e)
        {
            txt_a_q_1.Text = GetSelectedRadioButtonText(panel10);
            txt_a_q_2.Text = GetSelectedRadioButtonText(panel11);
            txt_a_q_3.Text = GetSelectedRadioButtonText(panel12);
            txt_a_q_4.Text = GetSelectedRadioButtonText(panel13);


            txt_amb_result.Text = Convert.ToString(Convert.ToInt32(txt_a_q_1.Text) + Convert.ToInt32(txt_a_q_2.Text) +
                Convert.ToInt32(txt_a_q_3.Text) + Convert.ToInt32(txt_a_q_4.Text));
            MessageBox.Show("Submitted successfuly");
            btn_refresh_2.Enabled = true;
            bt_amb_submit.Enabled = false;

        }
        private void Btn_ts_submit_Click(object sender, EventArgs e)
        {
            text_ts_1.Text = GetSelectedRadioButtonText(panel_1);
            text_ts_2.Text = GetSelectedRadioButtonText(panel_2);
            text_ts_3.Text = GetSelectedRadioButtonText(panel_3);
            text_ts_4.Text = GetSelectedRadioButtonText(panel_4);
            text_ts_5.Text = GetSelectedRadioButtonText(panel_5);
            text_ts_6.Text = GetSelectedRadioButtonText(panel_6);
            text_ts_7.Text = GetSelectedRadioButtonText(panel_7);
            text_ts_8.Text = GetSelectedRadioButtonText(panel_8);


            txt_ts_total.Text = Convert.ToString(Convert.ToInt32(text_ts_1.Text) + Convert.ToInt32(text_ts_2.Text) +
                Convert.ToInt32(text_ts_3.Text) + Convert.ToInt32(text_ts_4.Text) + Convert.ToInt32(text_ts_5.Text) + Convert.ToInt32(text_ts_6.Text) +
                Convert.ToInt32(text_ts_7.Text) + Convert.ToInt32(text_ts_8.Text));
            MessageBox.Show("Submitted successfuly");
            btn_refresh_3.Enabled = true;
            btn_ts_submit.Enabled = false;
        }

        private void Btn_refresh_Click(object sender, EventArgs e)
        {
            q1rb1.Checked = false;
            q1rb2.Checked = false;
            q1rb3.Checked = false;
            q1rb4.Checked = false;
            q1rb5.Checked = false;


            q2rb1.Checked = false;
            q2rb2.Checked = false;
            q2rb3.Checked = false;
            q2rb4.Checked = false;
            q2rb5.Checked = false;


            q3rb1.Checked = false;
            q3rb2.Checked = false;
            q3rb3.Checked = false;
            q3rb4.Checked = false;
            q3rb5.Checked = false;

            q4rb1.Checked = false;
            q4rb2.Checked = false;
            q4rb3.Checked = false;
            q4rb4.Checked = false;
            q4rb5.Checked = false;

            q5rb1.Checked = false;
            q5rb2.Checked = false;
            q5rb3.Checked = false;
            q5rb4.Checked = false;
            q5rb5.Checked = false;

            q6rb1.Checked = false;
            q6rb2.Checked = false;
            q6rb3.Checked = false;
            q6rb4.Checked = false;
            q6rb5.Checked = false;

            q7rb1.Checked = false;
            q7rb2.Checked = false;
            q7rb3.Checked = false;
            q7rb4.Checked = false;
            q7rb5.Checked = false;
            btn_submit.Enabled = false;
        }

        private void Btn_refresh_3_Click(object sender, EventArgs e)
        {
            tsq1rb1.Checked = false;
            tsq1rb2.Checked = false;
            tsq1rb3.Checked = false;
            tsq1rb4.Checked = false;
            tsq1rb5.Checked = false;


            tsq2rb1.Checked = false;
            tsq2rb2.Checked = false;
            tsq2rb3.Checked = false;
            tsq2rb4.Checked = false;
            tsq2rb5.Checked = false;


            tsq3rb1.Checked = false;
            tsq3rb2.Checked = false;
            tsq3rb3.Checked = false;
            tsq3rb4.Checked = false;
            tsq3rb5.Checked = false;

            tsq4rb1.Checked = false;
            tsq4rb2.Checked = false;
            tsq4rb3.Checked = false;
            tsq4rb4.Checked = false;
            tsq4rb5.Checked = false;

            tsq5rb1.Checked = false;
            tsq5rb2.Checked = false;
            tsq5rb3.Checked = false;
            tsq5rb4.Checked = false;
            tsq5rb5.Checked = false;

            tsq6rb1.Checked = false;
            tsq6rb2.Checked = false;
            tsq6rb3.Checked = false;
            tsq6rb4.Checked = false;
            tsq6rb5.Checked = false;

            tsq7rb1.Checked = false;
            tsq7rb2.Checked = false;
            tsq7rb3.Checked = false;
            tsq7rb4.Checked = false;
            tsq7rb5.Checked = false;

            tsq8rb1.Checked = false;
            tsq8rb2.Checked = false;
            tsq8rb3.Checked = false;
            tsq8rb4.Checked = false;
            tsq8rb5.Checked = false;
            btn_ts_submit.Enabled = false;
        }

        private void Btn_refresh_2_Click(object sender, EventArgs e)
        {
            aq1rb1.Checked = false;
            aq1rb1.Checked = false;
            aq1rb1.Checked = false;
            aq1rb1.Checked = false;
            aq1rb1.Checked = false;

            aq2rb1.Checked = false;
            aq2rb1.Checked = false;
            aq2rb1.Checked = false;
            aq2rb1.Checked = false;
            aq2rb1.Checked = false;

            aq3rb1.Checked = false;
            aq3rb1.Checked = false;
            aq3rb1.Checked = false;
            aq3rb1.Checked = false;
            aq3rb1.Checked = false;

            aq4rb1.Checked = false;
            aq4rb1.Checked = false;
            aq4rb1.Checked = false;
            aq4rb1.Checked = false;
            aq4rb1.Checked = false;
            bt_amb_submit.Enabled = false;
        }

        private void Button1_Click(object sender, EventArgs e)
        {

            if (
            q1rb1.Checked || q1rb2.Checked || q1rb3.Checked || q1rb4.Checked || q1rb5.Checked)
                if (
                q2rb1.Checked || q2rb2.Checked || q2rb3.Checked || q2rb4.Checked || q2rb5.Checked)
                    if (
                q3rb1.Checked ||
                q3rb2.Checked ||
                q3rb3.Checked ||
                q3rb4.Checked ||
                q3rb5.Checked)
                        if (
                q4rb1.Checked ||
                q4rb2.Checked ||
                q4rb3.Checked ||
                q4rb4.Checked ||
                q4rb5.Checked)
                            if (
                q5rb1.Checked ||
                q5rb2.Checked ||
                q5rb3.Checked ||
                q5rb4.Checked ||
                q5rb5.Checked)
                                if (
                q6rb1.Checked ||
                q6rb2.Checked ||
                q6rb3.Checked ||
                q6rb4.Checked ||
                q6rb5.Checked)
                                    if (
                q7rb1.Checked ||
                q7rb2.Checked ||
                q7rb3.Checked ||
                q7rb4.Checked ||
                q7rb5.Checked)
                                    {
                                        btn_submit.Enabled = true;
                                    }


        }

        private void Btn_chk_Click(object sender, EventArgs e)
        {
            if (aq1rb1.Checked || aq1rb2.Checked || aq1rb3.Checked || aq1rb4.Checked || aq1rb5.Checked)
                if (
               aq2rb1.Checked || aq2rb2.Checked || aq2rb3.Checked || aq2rb4.Checked || aq2rb5.Checked)
                    if (
            aq3rb1.Checked || aq3rb2.Checked || aq3rb3.Checked || aq3rb4.Checked || aq3rb5.Checked)
                        if (
                aq4rb1.Checked || aq4rb2.Checked || aq4rb3.Checked || aq4rb4.Checked || aq4rb5.Checked)
                        {
                            bt_amb_submit.Enabled = true;
                        }



        }

        private void Btn_check_Click(object sender, EventArgs e)
        {
            if (tsq1rb1.Checked || tsq1rb2.Checked || tsq1rb3.Checked || tsq1rb4.Checked || tsq1rb5.Checked)
                if (
            tsq2rb1.Checked || tsq2rb2.Checked || tsq2rb3.Checked || tsq2rb4.Checked || tsq2rb5.Checked)
                    if (
            tsq3rb1.Checked || tsq3rb2.Checked || tsq3rb3.Checked || tsq3rb4.Checked || tsq3rb5.Checked)
                        if (
            tsq4rb1.Checked || tsq4rb2.Checked || tsq4rb3.Checked || tsq4rb4.Checked || tsq4rb5.Checked)
                            if (
            tsq5rb1.Checked || tsq5rb2.Checked || tsq5rb3.Checked || tsq5rb4.Checked || tsq5rb5.Checked)
                                if (
            tsq6rb1.Checked || tsq6rb2.Checked || tsq6rb3.Checked || tsq6rb4.Checked || tsq6rb5.Checked)
                                    if (
            tsq7rb1.Checked || tsq7rb2.Checked || tsq7rb3.Checked || tsq7rb4.Checked || tsq7rb5.Checked)
                                        if (
            tsq8rb1.Checked || tsq8rb2.Checked || tsq8rb3.Checked || tsq8rb4.Checked || tsq8rb5.Checked)
                                        {
                                            btn_ts_submit.Enabled = true;
                                        }
        }

        private void Btn_evl1_Click(object sender, EventArgs e)
        {

            // int t = Convert.ToInt32(lbl_r_total.Text);
            if (lbl_r_total.Text == "")
            {
                MessageBox.Show("Please Insert Values");
                lbl_r_total.Focus();
                return;

            }
            if (Convert.ToInt32(lbl_r_total.Text) >= 7 && Convert.ToInt32(lbl_r_total.Text) <= 30)
            {
                txt_rec_status.Text = "Very low";
                txt_rec_msg.Text = @"LOW scorers deny that alcohol is causing them serious problems, 
reject diagnostic labels such as “problem drinker” and “alcoholic,” and do not express a desire for
change.";
            }
            else if (Convert.ToInt32(lbl_r_total.Text) == 31)
            {
                txt_rec_status.Text = "Above low";
            }
            else if (Convert.ToInt32(lbl_r_total.Text) >= 32 && Convert.ToInt32(lbl_r_total.Text) <= 33)

            {
                txt_rec_status.Text = "Medium";
            }
            else if (Convert.ToInt32(lbl_r_total.Text) == 34)
            {
                txt_rec_status.Text = "Medium +";
            }
            else if (Convert.ToInt32(lbl_r_total.Text) == 35)
            {
                txt_rec_status.Text = "High";
                txt_rec_msg.Text =
                @"HIGH scorers directly acknowledge that they are having problems related to their
drinking, tending to express a desire for change and to perceive that harm will continue if they 
do not change.";
            }
            else if (Convert.ToInt32(lbl_r_total.Text) > 35)
            {
                txt_rec_status.Text = "High +";
            }
        }

        private void Btn_evl3_Click(object sender, EventArgs e)
        {
            if (txt_ts_total.Text == "")
            {
                MessageBox.Show("Please Insert Values");
                txt_ts_total.Focus();
                return;

            }
            if (Convert.ToInt32(txt_ts_total.Text) >= 8 && Convert.ToInt32(txt_ts_total.Text) <= 30)
            {
                txt_ts_status.Text = "Very Low";
                txt_ts_msg.Text = @"LOW scorers report that they are not currently doing things to 
change their drinking, and have not made such changes recently.";
            }
            else if (Convert.ToInt32(txt_ts_total.Text) >= 31 && Convert.ToInt32(txt_ts_total.Text) <= 32)
            {
                txt_ts_status.Text = "Above Low";

            }
            else if (Convert.ToInt32(txt_ts_total.Text) == 33)
            {
                txt_ts_status.Text = "Medium";
            }
            else if (Convert.ToInt32(txt_ts_total.Text) >= 34 && Convert.ToInt32(txt_ts_total.Text) <= 35)
            {
                txt_ts_status.Text = "Medium +";

            }
            else if (Convert.ToInt32(txt_ts_total.Text) == 36)
            {
                txt_ts_status.Text = "High";
                txt_ts_msg.Text = @"HIGH scorers report that they are already doing things to 
make a positive change in their drinking, and may have experienced some success in this 
regard. Change is underway, and they may want help to persist or to prevent backsliding. 
A high score on this scale has been found to be predictive of successful change.";
            }
            else if (Convert.ToInt32(txt_ts_total.Text) >= 37 && Convert.ToInt32(txt_ts_total.Text) <= 38)
            {
                txt_ts_status.Text = "High +";
                txt_ts_msg.Text = @"HIGH scorers report that they are already doing things to make
a positive change in their drinking, and may have experienced some success in this regard.Change 
is underway, and they may want help to persist or to prevent backsliding.A high score on this 
scale has been found to be predictive of successful change.";
            }
            else if (Convert.ToInt32(txt_ts_total.Text) >= 39 && Convert.ToInt32(txt_ts_total.Text) <= 40)
            {
                txt_ts_status.Text = "Very High";
                txt_ts_msg.Text = @"HIGH scorers report that they are already doing things to make
a positive change in their drinking, and may have experienced some success in this regard.Change 
is underway, and they may want help to persist or to prevent backsliding.A high score on this 
scale has been found to be predictive of successful change.";

            }
        }

        private void Btn_evl2_Click(object sender, EventArgs e)
        {
            if (txt_amb_result.Text == "")
            {
                MessageBox.Show("Please Insert Values");
                lbl_r_total.Focus();
                return;

            }
            if (Convert.ToInt32(txt_amb_result.Text) >= 16 && Convert.ToInt32(txt_amb_result.Text) <= 20)
            {
                txt_am_status.Text = "High";
                txt_amb_msg.Text = @"HIGH scorers say that they sometimes wonder if they are in control
of their drinking, are drinking too much, are hurting other people, and/or are alcoholic. Thus a high
score reflects ambivalence or uncertainty. A high score here reflects some openness to reflection, as
might be particularly expected in the contemplation stage of change.";
            }

            else if (Convert.ToInt32(lbl_r_total.Text) >= 35 && Convert.ToInt32(txt_amb_result.Text) >= 4 && Convert.ToInt32(txt_amb_result.Text) <= 13)
            {
                txt_am_status.Text = "Low";
                txt_amb_msg.Text = @"Note that a person may score low on ambialence either because 
they “know” their drinking is causing problems (high Recognition).";
            }
            else if (Convert.ToInt32(lbl_r_total.Text) >= 7 && Convert.ToInt32(lbl_r_total.Text) <= 30 && Convert.ToInt32(txt_amb_result.Text) >= 4 && Convert.ToInt32(txt_amb_result.Text) <= 13)
            {
                txt_am_status.Text = "Low";
                txt_amb_msg.Text = @"because they “know” that they do not have drinking problems 
(low Recognition). Thus a low Ambivalence score should be interpreted in relation to the 
Recognition score.";
            }


        }

        private void Label80_Click(object sender, EventArgs e)
        {

        }

        private void PrintDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {

        }

        private void Button2_Click(object sender, EventArgs e)
        {

        }

        private void Btn_preview_Click(object sender, EventArgs e)
        {

            if (txt_rec_msg.Text != "")
            {
                String rec_msg = txt_rec_msg.Text;

                txt_rec_print.Text = rec_msg;
                btn_result_print.Enabled = true;
            }
            else
            {
                
                MessageBox.Show("Recognition Message is not filled.");
                btn_result_print.Enabled = false;
                return;
                
            }
            if (txt_amb_msg.Text != "")
            {
                String amb_msg = txt_amb_msg.Text;

                txt_amb_print.Text = amb_msg;
                btn_result_print.Enabled = true;

            }
            else
            {
                MessageBox.Show("Ambivalence Message is not filled.");
                btn_result_print.Enabled = false;
                return;
            }
            if (txt_ts_msg.Text != "")
            {
                String ts_msg = txt_ts_msg.Text;

                txt_ts_print.Text = ts_msg;
                btn_result_print.Enabled = true;
            }
            else
            {
                MessageBox.Show("Taking Step Message is not filled.");
                btn_result_print.Enabled = false;
                return;
            }

            if (txt_reg.Text != "")
            {
                
                btn_result_print.Enabled = true;
            }
            else
            {

                MessageBox.Show("Registration number is not filled.");
                btn_result_print.Enabled = false;
                return;

            }

            if (txt_pt_name.Text != "")
            {

                btn_result_print.Enabled = true;
            }
            else
            {

                MessageBox.Show("Patient Name is not Selected.");
                btn_result_print.Enabled = false;
                return;

            }

            if (rdbtn_female.Checked == true || rdbtn_male.Checked == true )
            {

                btn_result_print.Enabled = true;
            }
            else
            {

                MessageBox.Show("Gender is not Selected.");
                btn_result_print.Enabled = false;
                return;

            }

            if (txt_age.Text != "")
            {

                btn_result_print.Enabled = true;
            }
            else
            {

                MessageBox.Show("Age is not filled.");
                btn_result_print.Enabled = false;
                return;

            }


        }
      

        private void PrintDocument1_PrintPage_1(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {

           
        }
        public void writetoexcel()

        {
            

        }
        
        
        private void Btn_result_print_Click(object sender, EventArgs e)
        {

            string w = @Path.GetDirectoryName(Application.ExecutablePath).Trim() + "\\socrate_test.docx";


            word.Application wordApp = new word.Application();
            wordApp.Visible = true;

            word.Document document = wordApp.Documents.OpenNoRepairDialog(w);
           
            document.Activate();

            word.Table table1 = document.Tables[1];
            table1.Cell(1, 1).Range.Text = txt_reg.Text;

            word.Table table2 = document.Tables[2];
            table2.Cell(1, 1).Range.Text = txt_pt_name.Text;

            if(rdbtn_male.Checked==true)
            {
                word.Table table3 = document.Tables[3];
                table3.Cell(1, 1).Range.Text = rdbtn_male.Text;
            }
            else
            {
                word.Table table3 = document.Tables[3];
                table3.Cell(1, 1).Range.Text = rdbtn_female.Text;
            }

            word.Table table4 = document.Tables[4];
            table4.Cell(1, 1).Range.Text = txt_age.Text;


            word.Table table5 = document.Tables[5];
            table5.Cell(1,1).Range.Text = txt_rec_print.Text;

            word.Table table6 = document.Tables[6];
            table6.Cell(1, 1).Range.Text = txt_amb_print.Text;

            word.Table table7 = document.Tables[7];
            table7.Cell(1, 1).Range.Text = txt_ts_print.Text;

           // document.SaveAs2("C:\\Program Files (x86)\\Kaveri_IRCA\\Socrate_test_setup\\socrate_test.docx");
          //document.Close();
          //wordApp.Quit();
        }
    }
    }

