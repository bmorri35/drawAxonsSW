
namespace drawAxonsSW
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.b1 = new System.Windows.Forms.Button();
            this.b2 = new System.Windows.Forms.Button();
            this.b3 = new System.Windows.Forms.Button();
            this.b4 = new System.Windows.Forms.Button();
            this.b5 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // b1
            // 
            this.b1.Location = new System.Drawing.Point(12, 83);
            this.b1.Name = "b1";
            this.b1.Size = new System.Drawing.Size(86, 43);
            this.b1.TabIndex = 0;
            this.b1.Text = "Horizontal";
            this.b1.UseVisualStyleBackColor = true;
            this.b1.Click += new System.EventHandler(this.b1_Click);
            // 
            // b2
            // 
            this.b2.Location = new System.Drawing.Point(119, 83);
            this.b2.Name = "b2";
            this.b2.Size = new System.Drawing.Size(86, 43);
            this.b2.TabIndex = 1;
            this.b2.Text = "Posterior";
            this.b2.UseVisualStyleBackColor = true;
            this.b2.Click += new System.EventHandler(this.b2_Click);
            // 
            // b3
            // 
            this.b3.Location = new System.Drawing.Point(226, 83);
            this.b3.Name = "b3";
            this.b3.Size = new System.Drawing.Size(86, 43);
            this.b3.TabIndex = 2;
            this.b3.Text = "Superior";
            this.b3.UseVisualStyleBackColor = true;
            this.b3.Click += new System.EventHandler(this.b3_Click);
            // 
            // b4
            // 
            this.b4.Location = new System.Drawing.Point(333, 84);
            this.b4.Name = "b4";
            this.b4.Size = new System.Drawing.Size(86, 43);
            this.b4.TabIndex = 3;
            this.b4.Text = "Saccule";
            this.b4.UseVisualStyleBackColor = true;
            this.b4.Click += new System.EventHandler(this.b4_Click);
            // 
            // b5
            // 
            this.b5.Location = new System.Drawing.Point(440, 84);
            this.b5.Name = "b5";
            this.b5.Size = new System.Drawing.Size(86, 43);
            this.b5.TabIndex = 4;
            this.b5.Text = "Utricle";
            this.b5.UseVisualStyleBackColor = true;
            this.b5.Click += new System.EventHandler(this.b5_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(49, 33);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(441, 20);
            this.label1.TabIndex = 5;
            this.label1.Text = "Please choose the nerve branch you are generating axons for";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(539, 137);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.b5);
            this.Controls.Add(this.b4);
            this.Controls.Add(this.b3);
            this.Controls.Add(this.b2);
            this.Controls.Add(this.b1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button b1;
        private System.Windows.Forms.Button b2;
        private System.Windows.Forms.Button b3;
        private System.Windows.Forms.Button b4;
        private System.Windows.Forms.Button b5;
        private System.Windows.Forms.Label label1;
    }
}