namespace liquidacion
{
    partial class DetallePrecio
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
            this.lblTitulo = new System.Windows.Forms.Label();
            this.dtgDetalle = new System.Windows.Forms.DataGridView();
            this.btnSalir = new System.Windows.Forms.PictureBox();
            this.label1 = new System.Windows.Forms.Label();
            this.lblPrecioAnt = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.lblPrecioAct = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dtgDetalle)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnSalir)).BeginInit();
            this.SuspendLayout();
            // 
            // lblTitulo
            // 
            this.lblTitulo.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(150)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.lblTitulo.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitulo.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitulo.ForeColor = System.Drawing.Color.White;
            this.lblTitulo.Location = new System.Drawing.Point(0, 0);
            this.lblTitulo.Name = "lblTitulo";
            this.lblTitulo.Size = new System.Drawing.Size(875, 32);
            this.lblTitulo.TabIndex = 6;
            this.lblTitulo.Text = "DETALLE DE PRECIO DE VENTA POR FLETE";
            this.lblTitulo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.lblTitulo.Click += new System.EventHandler(this.lblTitulo_Click);
            this.lblTitulo.MouseDown += new System.Windows.Forms.MouseEventHandler(this.lblTitulo_MouseDown);
            // 
            // dtgDetalle
            // 
            this.dtgDetalle.AllowUserToAddRows = false;
            this.dtgDetalle.AllowUserToDeleteRows = false;
            this.dtgDetalle.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dtgDetalle.BackgroundColor = System.Drawing.Color.White;
            this.dtgDetalle.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dtgDetalle.GridColor = System.Drawing.Color.Black;
            this.dtgDetalle.Location = new System.Drawing.Point(12, 44);
            this.dtgDetalle.Name = "dtgDetalle";
            this.dtgDetalle.ReadOnly = true;
            this.dtgDetalle.Size = new System.Drawing.Size(851, 268);
            this.dtgDetalle.TabIndex = 7;
            // 
            // btnSalir
            // 
            this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(150)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.btnSalir.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnSalir.Image = global::liquidacion.Properties.Resources.eliminar;
            this.btnSalir.Location = new System.Drawing.Point(843, 2);
            this.btnSalir.Name = "btnSalir";
            this.btnSalir.Size = new System.Drawing.Size(30, 30);
            this.btnSalir.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.btnSalir.TabIndex = 415;
            this.btnSalir.TabStop = false;
            this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 21.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(112, 325);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(197, 40);
            this.label1.TabIndex = 416;
            this.label1.Text = "Precio factura:";
            // 
            // lblPrecioAnt
            // 
            this.lblPrecioAnt.AutoSize = true;
            this.lblPrecioAnt.Font = new System.Drawing.Font("Segoe UI", 21.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPrecioAnt.ForeColor = System.Drawing.Color.DarkOrange;
            this.lblPrecioAnt.Location = new System.Drawing.Point(315, 325);
            this.lblPrecioAnt.Name = "lblPrecioAnt";
            this.lblPrecioAnt.Size = new System.Drawing.Size(34, 40);
            this.lblPrecioAnt.TabIndex = 417;
            this.lblPrecioAnt.Text = "0";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Segoe UI", 21.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.White;
            this.label3.Location = new System.Drawing.Point(467, 325);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(229, 40);
            this.label3.TabIndex = 418;
            this.label3.Text = "Precio proforma:";
            // 
            // lblPrecioAct
            // 
            this.lblPrecioAct.AutoSize = true;
            this.lblPrecioAct.Font = new System.Drawing.Font("Segoe UI", 21.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPrecioAct.ForeColor = System.Drawing.Color.ForestGreen;
            this.lblPrecioAct.Location = new System.Drawing.Point(702, 325);
            this.lblPrecioAct.Name = "lblPrecioAct";
            this.lblPrecioAct.Size = new System.Drawing.Size(34, 40);
            this.lblPrecioAct.TabIndex = 419;
            this.lblPrecioAct.Text = "0";
            // 
            // DetallePrecio
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.DarkBlue;
            this.ClientSize = new System.Drawing.Size(875, 374);
            this.Controls.Add(this.lblPrecioAct);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.lblPrecioAnt);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnSalir);
            this.Controls.Add(this.dtgDetalle);
            this.Controls.Add(this.lblTitulo);
            this.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "DetallePrecio";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "DetallePrecio";
            ((System.ComponentModel.ISupportInitialize)(this.dtgDetalle)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnSalir)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblTitulo;
        private System.Windows.Forms.DataGridView dtgDetalle;
        private System.Windows.Forms.PictureBox btnSalir;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblPrecioAnt;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label lblPrecioAct;
    }
}