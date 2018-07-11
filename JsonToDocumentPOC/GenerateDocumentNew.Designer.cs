namespace JsonToDocumentPOC
{
    partial class GenerateDocumentNew
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
            this.label1 = new System.Windows.Forms.Label();
            this.txtJsonData = new System.Windows.Forms.RichTextBox();
            this.btnGenerateDocument = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(25, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(122, 17);
            this.label1.TabIndex = 0;
            this.label1.Text = "Json Data Input";
            // 
            // txtJsonData
            // 
            this.txtJsonData.Location = new System.Drawing.Point(28, 42);
            this.txtJsonData.Name = "txtJsonData";
            this.txtJsonData.Size = new System.Drawing.Size(1426, 501);
            this.txtJsonData.TabIndex = 2;
            this.txtJsonData.Text = "";
            // 
            // btnGenerateDocument
            // 
            this.btnGenerateDocument.Location = new System.Drawing.Point(645, 549);
            this.btnGenerateDocument.Name = "btnGenerateDocument";
            this.btnGenerateDocument.Size = new System.Drawing.Size(171, 45);
            this.btnGenerateDocument.TabIndex = 3;
            this.btnGenerateDocument.Text = "Generate Document";
            this.btnGenerateDocument.UseVisualStyleBackColor = true;
            this.btnGenerateDocument.Click += new System.EventHandler(this.btnGenerateDocument_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1484, 601);
            this.Controls.Add(this.btnGenerateDocument);
            this.Controls.Add(this.txtJsonData);
            this.Controls.Add(this.label1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "Form1";
            this.Text = "Json To Document";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.RichTextBox txtJsonData;
        private System.Windows.Forms.Button btnGenerateDocument;
    }
}

