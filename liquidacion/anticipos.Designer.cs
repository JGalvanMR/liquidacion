namespace liquidacion
{
    partial class anticipos
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.label1 = new System.Windows.Forms.Label();
            this.dtgMovimientos = new System.Windows.Forms.DataGridView();
            this.folio = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.nacexp = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tipo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.descuento = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnSalir = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.lblFecha = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.lblCultivo = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.lblAnticipo = new System.Windows.Forms.Label();
            this.lblFactura = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.lblContrato = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.lblCantidad = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.lblProveedor = new System.Windows.Forms.Label();
            this.lblSaldo = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.lblDescuento = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dtgMovimientos)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(217, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(152, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "MOVIMIENTOS DE ANTICIPO";
            // 
            // dtgMovimientos
            // 
            this.dtgMovimientos.AllowUserToAddRows = false;
            this.dtgMovimientos.AllowUserToDeleteRows = false;
            this.dtgMovimientos.BackgroundColor = System.Drawing.Color.White;
            this.dtgMovimientos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dtgMovimientos.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.folio,
            this.nacexp,
            this.tipo,
            this.descuento});
            this.dtgMovimientos.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.dtgMovimientos.GridColor = System.Drawing.Color.Black;
            this.dtgMovimientos.Location = new System.Drawing.Point(0, 154);
            this.dtgMovimientos.Name = "dtgMovimientos";
            this.dtgMovimientos.ReadOnly = true;
            this.dtgMovimientos.RowHeadersVisible = false;
            this.dtgMovimientos.Size = new System.Drawing.Size(584, 407);
            this.dtgMovimientos.TabIndex = 1;
            // 
            // folio
            // 
            this.folio.HeaderText = "LIQ / OC MP";
            this.folio.Name = "folio";
            this.folio.ReadOnly = true;
            this.folio.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // nacexp
            // 
            this.nacexp.HeaderText = "NAC / EXP";
            this.nacexp.Name = "nacexp";
            this.nacexp.ReadOnly = true;
            this.nacexp.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // tipo
            // 
            this.tipo.HeaderText = "TIPO";
            this.tipo.Name = "tipo";
            this.tipo.ReadOnly = true;
            this.tipo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // descuento
            // 
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            this.descuento.DefaultCellStyle = dataGridViewCellStyle1;
            this.descuento.HeaderText = "DESCUENTO";
            this.descuento.Name = "descuento";
            this.descuento.ReadOnly = true;
            this.descuento.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // btnSalir
            // 
            this.btnSalir.BackColor = System.Drawing.Color.Transparent;
            this.btnSalir.BackgroundImage = global::liquidacion.Properties.Resources.appbar_power___Copy;
            this.btnSalir.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.btnSalir.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.btnSalir.FlatAppearance.BorderSize = 2;
            this.btnSalir.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Black;
            this.btnSalir.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.btnSalir.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSalir.Font = new System.Drawing.Font("Lucida Sans Unicode", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSalir.ForeColor = System.Drawing.Color.Black;
            this.btnSalir.Location = new System.Drawing.Point(542, 12);
            this.btnSalir.Name = "btnSalir";
            this.btnSalir.Size = new System.Drawing.Size(30, 30);
            this.btnSalir.TabIndex = 395;
            this.btnSalir.UseVisualStyleBackColor = false;
            this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(69, 73);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(40, 13);
            this.label2.TabIndex = 396;
            this.label2.Text = "Fecha:";
            // 
            // lblFecha
            // 
            this.lblFecha.AutoSize = true;
            this.lblFecha.BackColor = System.Drawing.Color.Transparent;
            this.lblFecha.ForeColor = System.Drawing.Color.White;
            this.lblFecha.Location = new System.Drawing.Point(115, 73);
            this.lblFecha.Name = "lblFecha";
            this.lblFecha.Size = new System.Drawing.Size(11, 13);
            this.lblFecha.TabIndex = 397;
            this.lblFecha.Text = "-";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.BackColor = System.Drawing.Color.Transparent;
            this.label4.ForeColor = System.Drawing.Color.White;
            this.label4.Location = new System.Drawing.Point(195, 73);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(46, 13);
            this.label4.TabIndex = 398;
            this.label4.Text = "Cultivo:";
            // 
            // lblCultivo
            // 
            this.lblCultivo.AutoSize = true;
            this.lblCultivo.BackColor = System.Drawing.Color.Transparent;
            this.lblCultivo.ForeColor = System.Drawing.Color.White;
            this.lblCultivo.Location = new System.Drawing.Point(247, 73);
            this.lblCultivo.Name = "lblCultivo";
            this.lblCultivo.Size = new System.Drawing.Size(11, 13);
            this.lblCultivo.TabIndex = 399;
            this.lblCultivo.Text = "-";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.BackColor = System.Drawing.Color.Transparent;
            this.label6.ForeColor = System.Drawing.Color.White;
            this.label6.Location = new System.Drawing.Point(354, 73);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(53, 13);
            this.label6.TabIndex = 400;
            this.label6.Text = "Anticipo:";
            // 
            // lblAnticipo
            // 
            this.lblAnticipo.AutoSize = true;
            this.lblAnticipo.BackColor = System.Drawing.Color.Transparent;
            this.lblAnticipo.ForeColor = System.Drawing.Color.White;
            this.lblAnticipo.Location = new System.Drawing.Point(413, 73);
            this.lblAnticipo.Name = "lblAnticipo";
            this.lblAnticipo.Size = new System.Drawing.Size(11, 13);
            this.lblAnticipo.TabIndex = 401;
            this.lblAnticipo.Text = "-";
            // 
            // lblFactura
            // 
            this.lblFactura.AutoSize = true;
            this.lblFactura.BackColor = System.Drawing.Color.Transparent;
            this.lblFactura.ForeColor = System.Drawing.Color.White;
            this.lblFactura.Location = new System.Drawing.Point(282, 100);
            this.lblFactura.Name = "lblFactura";
            this.lblFactura.Size = new System.Drawing.Size(11, 13);
            this.lblFactura.TabIndex = 407;
            this.lblFactura.Text = "-";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.BackColor = System.Drawing.Color.Transparent;
            this.label9.ForeColor = System.Drawing.Color.White;
            this.label9.Location = new System.Drawing.Point(228, 100);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(48, 13);
            this.label9.TabIndex = 406;
            this.label9.Text = "Factura:";
            // 
            // lblContrato
            // 
            this.lblContrato.AutoSize = true;
            this.lblContrato.BackColor = System.Drawing.Color.Transparent;
            this.lblContrato.ForeColor = System.Drawing.Color.White;
            this.lblContrato.Location = new System.Drawing.Point(131, 100);
            this.lblContrato.Name = "lblContrato";
            this.lblContrato.Size = new System.Drawing.Size(11, 13);
            this.lblContrato.TabIndex = 405;
            this.lblContrato.Text = "-";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.BackColor = System.Drawing.Color.Transparent;
            this.label11.ForeColor = System.Drawing.Color.White;
            this.label11.Location = new System.Drawing.Point(69, 100);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(56, 13);
            this.label11.TabIndex = 404;
            this.label11.Text = "Contrato:";
            // 
            // lblCantidad
            // 
            this.lblCantidad.AutoSize = true;
            this.lblCantidad.BackColor = System.Drawing.Color.Transparent;
            this.lblCantidad.ForeColor = System.Drawing.Color.White;
            this.lblCantidad.Location = new System.Drawing.Point(132, 128);
            this.lblCantidad.Name = "lblCantidad";
            this.lblCantidad.Size = new System.Drawing.Size(11, 13);
            this.lblCantidad.TabIndex = 403;
            this.lblCantidad.Text = "-";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.BackColor = System.Drawing.Color.Transparent;
            this.label13.ForeColor = System.Drawing.Color.White;
            this.label13.Location = new System.Drawing.Point(69, 128);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(57, 13);
            this.label13.TabIndex = 402;
            this.label13.Text = "Cantidad:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.BackColor = System.Drawing.Color.Transparent;
            this.label5.ForeColor = System.Drawing.Color.White;
            this.label5.Location = new System.Drawing.Point(69, 44);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(62, 13);
            this.label5.TabIndex = 408;
            this.label5.Text = "Proveedor:";
            // 
            // lblProveedor
            // 
            this.lblProveedor.AutoSize = true;
            this.lblProveedor.BackColor = System.Drawing.Color.Transparent;
            this.lblProveedor.ForeColor = System.Drawing.Color.White;
            this.lblProveedor.Location = new System.Drawing.Point(137, 44);
            this.lblProveedor.Name = "lblProveedor";
            this.lblProveedor.Size = new System.Drawing.Size(11, 13);
            this.lblProveedor.TabIndex = 409;
            this.lblProveedor.Text = "-";
            // 
            // lblSaldo
            // 
            this.lblSaldo.AutoSize = true;
            this.lblSaldo.BackColor = System.Drawing.Color.Transparent;
            this.lblSaldo.ForeColor = System.Drawing.Color.White;
            this.lblSaldo.Location = new System.Drawing.Point(443, 128);
            this.lblSaldo.Name = "lblSaldo";
            this.lblSaldo.Size = new System.Drawing.Size(11, 13);
            this.lblSaldo.TabIndex = 411;
            this.lblSaldo.Text = "-";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.BackColor = System.Drawing.Color.Transparent;
            this.label7.ForeColor = System.Drawing.Color.White;
            this.label7.Location = new System.Drawing.Point(398, 128);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(39, 13);
            this.label7.TabIndex = 410;
            this.label7.Text = "Saldo:";
            // 
            // lblDescuento
            // 
            this.lblDescuento.AutoSize = true;
            this.lblDescuento.BackColor = System.Drawing.Color.Transparent;
            this.lblDescuento.ForeColor = System.Drawing.Color.White;
            this.lblDescuento.Location = new System.Drawing.Point(301, 128);
            this.lblDescuento.Name = "lblDescuento";
            this.lblDescuento.Size = new System.Drawing.Size(11, 13);
            this.lblDescuento.TabIndex = 413;
            this.lblDescuento.Text = "-";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.BackColor = System.Drawing.Color.Transparent;
            this.label8.ForeColor = System.Drawing.Color.White;
            this.label8.Location = new System.Drawing.Point(228, 128);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(65, 13);
            this.label8.TabIndex = 412;
            this.label8.Text = "Descuento:";
            // 
            // anticipos
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlDark;
            this.ClientSize = new System.Drawing.Size(584, 561);
            this.ControlBox = false;
            this.Controls.Add(this.lblDescuento);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.lblSaldo);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.lblProveedor);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.lblFactura);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.lblContrato);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.lblCantidad);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.lblAnticipo);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.lblCultivo);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.lblFecha);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnSalir);
            this.Controls.Add(this.dtgMovimientos);
            this.Controls.Add(this.label1);
            this.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "anticipos";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Anticipo";
            this.Load += new System.EventHandler(this.anticipos_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dtgMovimientos)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGridView dtgMovimientos;
        private System.Windows.Forms.Button btnSalir;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label lblFecha;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label lblCultivo;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label lblAnticipo;
        private System.Windows.Forms.Label lblFactura;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label lblContrato;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label lblCantidad;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label lblProveedor;
        private System.Windows.Forms.Label lblSaldo;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label lblDescuento;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.DataGridViewTextBoxColumn folio;
        private System.Windows.Forms.DataGridViewTextBoxColumn nacexp;
        private System.Windows.Forms.DataGridViewTextBoxColumn tipo;
        private System.Windows.Forms.DataGridViewTextBoxColumn descuento;
    }
}