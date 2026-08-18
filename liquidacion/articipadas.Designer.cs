namespace liquidacion
{
    partial class articipadas
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            this.label1 = new System.Windows.Forms.Label();
            this.lblClave = new System.Windows.Forms.Label();
            this.lblNombre = new System.Windows.Forms.Label();
            this.dtgOrdenes = new System.Windows.Forms.DataGridView();
            this.oc = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fecha = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cve_prod = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.nom_prod = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cajas = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.precio = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.importe = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.conse = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.lblProveedor = new System.Windows.Forms.Label();
            this.lblCveProv = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.lblCantidad = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.dtgLiquidaciones = new System.Windows.Forms.DataGridView();
            this.btnAceptar = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.lblCantLiquidado = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.lblDiferencia = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dtgOrdenes)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dtgLiquidaciones)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(11, 46);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(59, 15);
            this.label1.TabIndex = 0;
            this.label1.Text = "Producto:";
            // 
            // lblClave
            // 
            this.lblClave.AutoSize = true;
            this.lblClave.ForeColor = System.Drawing.Color.White;
            this.lblClave.Location = new System.Drawing.Point(83, 46);
            this.lblClave.Name = "lblClave";
            this.lblClave.Size = new System.Drawing.Size(13, 15);
            this.lblClave.TabIndex = 1;
            this.lblClave.Text = "-";
            // 
            // lblNombre
            // 
            this.lblNombre.AutoSize = true;
            this.lblNombre.ForeColor = System.Drawing.Color.White;
            this.lblNombre.Location = new System.Drawing.Point(181, 46);
            this.lblNombre.Name = "lblNombre";
            this.lblNombre.Size = new System.Drawing.Size(13, 15);
            this.lblNombre.TabIndex = 2;
            this.lblNombre.Text = "-";
            // 
            // dtgOrdenes
            // 
            this.dtgOrdenes.AllowUserToAddRows = false;
            this.dtgOrdenes.AllowUserToDeleteRows = false;
            this.dtgOrdenes.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dtgOrdenes.BackgroundColor = System.Drawing.Color.White;
            this.dtgOrdenes.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dtgOrdenes.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.oc,
            this.fecha,
            this.cve_prod,
            this.nom_prod,
            this.cajas,
            this.precio,
            this.importe,
            this.conse});
            this.dtgOrdenes.GridColor = System.Drawing.Color.Black;
            this.dtgOrdenes.Location = new System.Drawing.Point(14, 79);
            this.dtgOrdenes.Name = "dtgOrdenes";
            this.dtgOrdenes.ReadOnly = true;
            this.dtgOrdenes.RowHeadersVisible = false;
            this.dtgOrdenes.Size = new System.Drawing.Size(793, 232);
            this.dtgOrdenes.TabIndex = 3;
            this.dtgOrdenes.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dtgOrdenes_CellDoubleClick);
            // 
            // oc
            // 
            this.oc.HeaderText = "OC";
            this.oc.Name = "oc";
            this.oc.ReadOnly = true;
            this.oc.Width = 49;
            // 
            // fecha
            // 
            this.fecha.HeaderText = "Fecha";
            this.fecha.Name = "fecha";
            this.fecha.ReadOnly = true;
            this.fecha.Width = 63;
            // 
            // cve_prod
            // 
            this.cve_prod.HeaderText = "Cve Prod";
            this.cve_prod.Name = "cve_prod";
            this.cve_prod.ReadOnly = true;
            this.cve_prod.Width = 80;
            // 
            // nom_prod
            // 
            this.nom_prod.HeaderText = "Nom Prod";
            this.nom_prod.Name = "nom_prod";
            this.nom_prod.ReadOnly = true;
            this.nom_prod.Width = 86;
            // 
            // cajas
            // 
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            this.cajas.DefaultCellStyle = dataGridViewCellStyle1;
            this.cajas.HeaderText = "Cajas";
            this.cajas.Name = "cajas";
            this.cajas.ReadOnly = true;
            this.cajas.Width = 61;
            // 
            // precio
            // 
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            this.precio.DefaultCellStyle = dataGridViewCellStyle2;
            this.precio.HeaderText = "Precio";
            this.precio.Name = "precio";
            this.precio.ReadOnly = true;
            this.precio.Width = 65;
            // 
            // importe
            // 
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            this.importe.DefaultCellStyle = dataGridViewCellStyle3;
            this.importe.HeaderText = "Importe";
            this.importe.Name = "importe";
            this.importe.ReadOnly = true;
            this.importe.Width = 75;
            // 
            // conse
            // 
            this.conse.HeaderText = "Conse";
            this.conse.Name = "conse";
            this.conse.ReadOnly = true;
            this.conse.Width = 66;
            // 
            // lblProveedor
            // 
            this.lblProveedor.AutoSize = true;
            this.lblProveedor.ForeColor = System.Drawing.Color.White;
            this.lblProveedor.Location = new System.Drawing.Point(181, 19);
            this.lblProveedor.Name = "lblProveedor";
            this.lblProveedor.Size = new System.Drawing.Size(13, 15);
            this.lblProveedor.TabIndex = 6;
            this.lblProveedor.Text = "-";
            // 
            // lblCveProv
            // 
            this.lblCveProv.AutoSize = true;
            this.lblCveProv.ForeColor = System.Drawing.Color.White;
            this.lblCveProv.Location = new System.Drawing.Point(83, 19);
            this.lblCveProv.Name = "lblCveProv";
            this.lblCveProv.Size = new System.Drawing.Size(13, 15);
            this.lblCveProv.TabIndex = 5;
            this.lblCveProv.Text = "-";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.ForeColor = System.Drawing.Color.White;
            this.label4.Location = new System.Drawing.Point(12, 19);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(65, 15);
            this.label4.TabIndex = 4;
            this.label4.Text = "Proveedor:";
            // 
            // lblCantidad
            // 
            this.lblCantidad.AutoSize = true;
            this.lblCantidad.ForeColor = System.Drawing.Color.White;
            this.lblCantidad.Location = new System.Drawing.Point(501, 19);
            this.lblCantidad.Name = "lblCantidad";
            this.lblCantidad.Size = new System.Drawing.Size(13, 15);
            this.lblCantidad.TabIndex = 346;
            this.lblCantidad.Text = "-";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.ForeColor = System.Drawing.Color.White;
            this.label6.Location = new System.Drawing.Point(384, 19);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(111, 15);
            this.label6.TabIndex = 345;
            this.label6.Text = "Cantidad a liquidar:";
            // 
            // dtgLiquidaciones
            // 
            this.dtgLiquidaciones.AllowUserToAddRows = false;
            this.dtgLiquidaciones.AllowUserToDeleteRows = false;
            this.dtgLiquidaciones.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dtgLiquidaciones.BackgroundColor = System.Drawing.Color.White;
            this.dtgLiquidaciones.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dtgLiquidaciones.GridColor = System.Drawing.Color.Black;
            this.dtgLiquidaciones.Location = new System.Drawing.Point(14, 317);
            this.dtgLiquidaciones.Name = "dtgLiquidaciones";
            this.dtgLiquidaciones.ReadOnly = true;
            this.dtgLiquidaciones.RowHeadersVisible = false;
            this.dtgLiquidaciones.Size = new System.Drawing.Size(793, 181);
            this.dtgLiquidaciones.TabIndex = 347;
            // 
            // btnAceptar
            // 
            this.btnAceptar.BackColor = System.Drawing.Color.Transparent;
            this.btnAceptar.BackgroundImage = global::liquidacion.Properties.Resources.appbar_check;
            this.btnAceptar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.btnAceptar.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.btnAceptar.FlatAppearance.BorderSize = 2;
            this.btnAceptar.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Black;
            this.btnAceptar.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.btnAceptar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAceptar.Font = new System.Drawing.Font("Lucida Sans Unicode", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAceptar.ForeColor = System.Drawing.Color.Black;
            this.btnAceptar.Location = new System.Drawing.Point(385, 514);
            this.btnAceptar.Name = "btnAceptar";
            this.btnAceptar.Size = new System.Drawing.Size(50, 50);
            this.btnAceptar.TabIndex = 348;
            this.btnAceptar.UseVisualStyleBackColor = false;
            this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(384, 46);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(111, 15);
            this.label2.TabIndex = 349;
            this.label2.Text = "Cantidad liquidado:";
            // 
            // lblCantLiquidado
            // 
            this.lblCantLiquidado.AutoSize = true;
            this.lblCantLiquidado.ForeColor = System.Drawing.Color.White;
            this.lblCantLiquidado.Location = new System.Drawing.Point(501, 46);
            this.lblCantLiquidado.Name = "lblCantLiquidado";
            this.lblCantLiquidado.Size = new System.Drawing.Size(13, 15);
            this.lblCantLiquidado.TabIndex = 350;
            this.lblCantLiquidado.Text = "-";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.ForeColor = System.Drawing.Color.White;
            this.label3.Location = new System.Drawing.Point(590, 46);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(64, 15);
            this.label3.TabIndex = 351;
            this.label3.Text = "Diferencia:";
            // 
            // lblDiferencia
            // 
            this.lblDiferencia.AutoSize = true;
            this.lblDiferencia.ForeColor = System.Drawing.Color.White;
            this.lblDiferencia.Location = new System.Drawing.Point(660, 46);
            this.lblDiferencia.Name = "lblDiferencia";
            this.lblDiferencia.Size = new System.Drawing.Size(13, 15);
            this.lblDiferencia.TabIndex = 352;
            this.lblDiferencia.Text = "-";
            // 
            // articipadas
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.DarkBlue;
            this.ClientSize = new System.Drawing.Size(820, 579);
            this.Controls.Add(this.lblDiferencia);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.lblCantLiquidado);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnAceptar);
            this.Controls.Add(this.dtgLiquidaciones);
            this.Controls.Add(this.lblCantidad);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.lblProveedor);
            this.Controls.Add(this.lblCveProv);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.dtgOrdenes);
            this.Controls.Add(this.lblNombre);
            this.Controls.Add(this.lblClave);
            this.Controls.Add(this.label1);
            this.Font = new System.Drawing.Font("Lucida Sans Unicode", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "articipadas";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Ordenes Anticipadas";
            this.Load += new System.EventHandler(this.articipadas_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dtgOrdenes)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dtgLiquidaciones)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblClave;
        private System.Windows.Forms.Label lblNombre;
        private System.Windows.Forms.DataGridView dtgOrdenes;
        private System.Windows.Forms.Label lblProveedor;
        private System.Windows.Forms.Label lblCveProv;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label lblCantidad;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.DataGridViewTextBoxColumn oc;
        private System.Windows.Forms.DataGridViewTextBoxColumn fecha;
        private System.Windows.Forms.DataGridViewTextBoxColumn cve_prod;
        private System.Windows.Forms.DataGridViewTextBoxColumn nom_prod;
        private System.Windows.Forms.DataGridViewTextBoxColumn cajas;
        private System.Windows.Forms.DataGridViewTextBoxColumn precio;
        private System.Windows.Forms.DataGridViewTextBoxColumn importe;
        private System.Windows.Forms.DataGridViewTextBoxColumn conse;
        private System.Windows.Forms.DataGridView dtgLiquidaciones;
        private System.Windows.Forms.Button btnAceptar;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label lblCantLiquidado;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label lblDiferencia;
    }
}