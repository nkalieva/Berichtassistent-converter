using System;

namespace Resave
{
    partial class Form1
    {
        /// <summary>
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
        /// <param name="disposing">True, wenn verwaltete Ressourcen gelöscht werden sollen; andernfalls False.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Vom Windows Form-Designer generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            btnShow = new System.Windows.Forms.Button();
            btnSave = new System.Windows.Forms.Button();
            progressBar1 = new System.Windows.Forms.ProgressBar();
            lblProgress = new System.Windows.Forms.Label();
            lblText = new System.Windows.Forms.Label();
            lblProgress1 = new System.Windows.Forms.Label();
            progressBar2 = new System.Windows.Forms.ProgressBar();
            lblText1 = new System.Windows.Forms.Label();
            lblProgress2 = new System.Windows.Forms.Label();
            label1 = new System.Windows.Forms.Label();
            SuspendLayout();
            // 
            // btnShow
            // 
            btnShow.Location = new System.Drawing.Point(66, 22);
            btnShow.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            btnShow.Name = "btnShow";
            btnShow.Size = new System.Drawing.Size(271, 39);
            btnShow.TabIndex = 0;
            btnShow.Text = "Zu konvertierende Excel-Datei öffnen";
            btnShow.UseVisualStyleBackColor = true;
            btnShow.Click += btnShow_Click;
            // 
            // btnSave
            // 
            btnSave.BackColor = System.Drawing.SystemColors.Control;
            btnSave.Enabled = false;
            btnSave.Location = new System.Drawing.Point(366, 22);
            btnSave.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            btnSave.Name = "btnSave";
            btnSave.Size = new System.Drawing.Size(141, 39);
            btnSave.TabIndex = 6;
            btnSave.Text = "Speichern";
            btnSave.UseVisualStyleBackColor = false;
            btnSave.Click += btnSave_Click;
            // 
            // progressBar1
            // 
            progressBar1.Location = new System.Drawing.Point(66, 101);
            progressBar1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            progressBar1.Name = "progressBar1";
            progressBar1.Size = new System.Drawing.Size(441, 11);
            progressBar1.Style = System.Windows.Forms.ProgressBarStyle.Marquee;
            progressBar1.TabIndex = 7;
            progressBar1.Visible = false;
            progressBar1.Click += progressBar2_Click;
            // 
            // lblProgress
            // 
            lblProgress.AutoSize = true;
            lblProgress.Location = new System.Drawing.Point(507, 88);
            lblProgress.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            lblProgress.Name = "lblProgress";
            lblProgress.Size = new System.Drawing.Size(0, 15);
            lblProgress.TabIndex = 8;
            lblProgress.Visible = false;
            // 
            // lblText
            // 
            lblText.AllowDrop = true;
            lblText.ImageAlign = System.Drawing.ContentAlignment.TopLeft;
            lblText.Location = new System.Drawing.Point(63, 79);
            lblText.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            lblText.Name = "lblText";
            lblText.RightToLeft = System.Windows.Forms.RightToLeft.No;
            lblText.Size = new System.Drawing.Size(353, 15);
            lblText.TabIndex = 9;
            lblText.Text = "Datei wird verarbeitet...";
            lblText.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            lblText.Visible = false;
            lblText.Click += lblText_Click;
            // 
            // lblProgress1
            // 
            lblProgress1.Location = new System.Drawing.Point(363, 79);
            lblProgress1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            lblProgress1.Name = "lblProgress1";
            lblProgress1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            lblProgress1.Size = new System.Drawing.Size(141, 15);
            lblProgress1.TabIndex = 10;
            lblProgress1.Text = "0";
            lblProgress1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            lblProgress1.Visible = false;
            lblProgress1.Click += lblProgress1_Click;
            // 
            // progressBar2
            // 
            progressBar2.Location = new System.Drawing.Point(66, 148);
            progressBar2.Name = "progressBar2";
            progressBar2.Size = new System.Drawing.Size(441, 10);
            progressBar2.TabIndex = 11;
            progressBar2.Visible = false;
            progressBar2.Click += progressBar2_Click;
            // 
            // lblText1
            // 
            lblText1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            lblText1.Location = new System.Drawing.Point(63, 130);
            lblText1.Name = "lblText1";
            lblText1.Size = new System.Drawing.Size(297, 15);
            lblText1.TabIndex = 12;
            lblText1.Text = "Aktuelles Tabellenblatt:";
            lblText1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            lblText1.Visible = false;
            lblText1.Click += lblText1_Click;
            // 
            // lblProgress2
            // 
            lblProgress2.ForeColor = System.Drawing.SystemColors.ControlText;
            lblProgress2.ImageAlign = System.Drawing.ContentAlignment.BottomRight;
            lblProgress2.Location = new System.Drawing.Point(366, 130);
            lblProgress2.Name = "lblProgress2";
            lblProgress2.RightToLeft = System.Windows.Forms.RightToLeft.No;
            lblProgress2.Size = new System.Drawing.Size(141, 15);
            lblProgress2.TabIndex = 13;
            lblProgress2.Text = "leer";
            lblProgress2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            lblProgress2.Visible = false;
            lblProgress2.Click += lblProgress2_Click;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new System.Drawing.Point(189, 94);
            label1.Name = "label1";
            label1.Size = new System.Drawing.Size(171, 15);
            label1.TabIndex = 14;
            label1.Text = "Keine Kommentare gefunden...";
            label1.Visible = false;
            // 
            // Form1
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(590, 191);
            Controls.Add(label1);
            Controls.Add(lblProgress2);
            Controls.Add(lblText1);
            Controls.Add(progressBar2);
            Controls.Add(lblProgress1);
            Controls.Add(lblText);
            Controls.Add(lblProgress);
            Controls.Add(progressBar1);
            Controls.Add(btnSave);
            Controls.Add(btnShow);
            Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            Name = "Form1";
            Text = "Form1";
            Load += Form1_Load_1;
            ResumeLayout(false);
            PerformLayout();
        }

        private void lblProgress1_Click(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }

        private void lblText_Click(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }

        #endregion

        private System.Windows.Forms.Button btnShow;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label lblProgress;
        private System.Windows.Forms.Label lblText;
        private System.Windows.Forms.Label lblProgress1;
        private System.Windows.Forms.ProgressBar progressBar2;
        private System.Windows.Forms.Label lblText1;
        private System.Windows.Forms.Label lblProgress2;
        private System.Windows.Forms.Label label1;
    }
}

