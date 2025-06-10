namespace MsTool
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            button1 = new Button();
            button2 = new Button();
            label1 = new Label();
            label2 = new Label();
            button3 = new Button();
            checkBox1 = new CheckBox();
            AssumptionsCB = new CheckBox();
            AnalyticsCB = new CheckBox();
            SuspendLayout();
            // 
            // button1
            // 
            button1.Location = new Point(125, 88);
            button1.Name = "button1";
            button1.Size = new Size(178, 101);
            button1.TabIndex = 0;
            button1.Text = "Moj fajl";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // button2
            // 
            button2.Location = new Point(483, 88);
            button2.Name = "button2";
            button2.Size = new Size(178, 101);
            button2.TabIndex = 1;
            button2.Text = "Fajl poreske uprave";
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(125, 70);
            label1.Name = "label1";
            label1.Size = new Size(0, 15);
            label1.TabIndex = 2;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(486, 70);
            label2.Name = "label2";
            label2.Size = new Size(0, 15);
            label2.TabIndex = 3;
            // 
            // button3
            // 
            button3.Location = new Point(308, 285);
            button3.Name = "button3";
            button3.Size = new Size(178, 101);
            button3.TabIndex = 6;
            button3.Text = "Uporedi";
            button3.UseVisualStyleBackColor = true;
            button3.Click += button3_Click;
            // 
            // checkBox1
            // 
            checkBox1.AutoSize = true;
            checkBox1.Location = new Point(333, 128);
            checkBox1.Name = "checkBox1";
            checkBox1.Size = new Size(81, 19);
            checkBox1.TabIndex = 7;
            checkBox1.Text = "Sve greske";
            checkBox1.UseVisualStyleBackColor = true;
            // 
            // AssumptionsCB
            // 
            AssumptionsCB.AutoSize = true;
            AssumptionsCB.Location = new Point(333, 170);
            AssumptionsCB.Name = "AssumptionsCB";
            AssumptionsCB.Size = new Size(131, 19);
            AssumptionsCB.TabIndex = 8;
            AssumptionsCB.Text = "Prikazi pretpostavke";
            AssumptionsCB.UseVisualStyleBackColor = true;
            // 
            // AnalyticsCB
            // 
            AnalyticsCB.AutoSize = true;
            AnalyticsCB.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            AnalyticsCB.Location = new Point(333, 87);
            AnalyticsCB.Name = "AnalyticsCB";
            AnalyticsCB.Size = new Size(123, 19);
            AnalyticsCB.TabIndex = 9;
            AnalyticsCB.Text = "Analiticke kartice";
            AnalyticsCB.UseVisualStyleBackColor = true;
            AnalyticsCB.CheckedChanged += AnalyticsCB_CheckedChanged;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(AnalyticsCB);
            Controls.Add(AssumptionsCB);
            Controls.Add(checkBox1);
            Controls.Add(button3);
            Controls.Add(label2);
            Controls.Add(label1);
            Controls.Add(button2);
            Controls.Add(button1);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "Form1";
            Text = "MsTool v2.3";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button button1;
        private Button button2;
        private Label label1;
        private Label label2;
        private Button button3;
        private CheckBox checkBox1;
        private CheckBox AssumptionsCB;
        private CheckBox AnalyticsCB;
    }
}
