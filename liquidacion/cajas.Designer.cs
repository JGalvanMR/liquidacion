namespace liquidacion
{
    partial class cajas
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(cajas));
            this.txtCantidad = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.btnCierra = new System.Windows.Forms.Button();
            this.btnCantidad = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // txtCantidad
            // 
            this.txtCantidad.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtCantidad.Font = new System.Drawing.Font("Lucida Sans Unicode", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtCantidad.Location = new System.Drawing.Point(109, 46);
            this.txtCantidad.Name = "txtCantidad";
            this.txtCantidad.Size = new System.Drawing.Size(100, 26);
            this.txtCantidad.TabIndex = 376;
            this.txtCantidad.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtCantidad.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCantidad_KeyPress);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.BackColor = System.Drawing.Color.Transparent;
            this.label10.Font = new System.Drawing.Font("Lucida Sans Unicode", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.ForeColor = System.Drawing.Color.White;
            this.label10.Location = new System.Drawing.Point(22, 18);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(122, 16);
            this.label10.TabIndex = 374;
            this.label10.Text = "Cantidad a liquidar:";
            // 
            // btnCierra
            // 
            this.btnCierra.BackColor = System.Drawing.Color.Transparent;
            this.btnCierra.BackgroundImage = global::liquidacion.Properties.Resources.appbar_power___Copy;
            this.btnCierra.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.btnCierra.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.btnCierra.FlatAppearance.BorderSize = 2;
            this.btnCierra.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Black;
            this.btnCierra.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.btnCierra.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCierra.Font = new System.Drawing.Font("Lucida Sans Unicode", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCierra.ForeColor = System.Drawing.Color.Black;
            this.btnCierra.Location = new System.Drawing.Point(266, 11);
            this.btnCierra.Name = "btnCierra";
            this.btnCierra.Size = new System.Drawing.Size(30, 30);
            this.btnCierra.TabIndex = 377;
            this.btnCierra.UseVisualStyleBackColor = false;
            this.btnCierra.Click += new System.EventHandler(this.btnCierra_Click);
            // 
            // btnCantidad
            // 
            this.btnCantidad.BackColor = System.Drawing.Color.Transparent;
            this.btnCantidad.BackgroundImage = global::liquidacion.Properties.Resources.appbar_check;
            this.btnCantidad.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.btnCantidad.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.btnCantidad.FlatAppearance.BorderSize = 2;
            this.btnCantidad.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Black;
            this.btnCantidad.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.btnCantidad.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCantidad.Font = new System.Drawing.Font("Lucida Sans Unicode", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCantidad.ForeColor = System.Drawing.Color.Black;
            this.btnCantidad.Location = new System.Drawing.Point(266, 57);
            this.btnCantidad.Name = "btnCantidad";
            this.btnCantidad.Size = new System.Drawing.Size(30, 30);
            this.btnCantidad.TabIndex = 375;
            this.btnCantidad.UseVisualStyleBackColor = false;
            this.btnCantidad.Click += new System.EventHandler(this.btnCantidad_Click);
            // 
            // cajas
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.DarkBlue;
            this.ClientSize = new System.Drawing.Size(322, 104);
            this.ControlBox = false;
            this.Controls.Add(this.btnCierra);
            this.Controls.Add(this.txtCantidad);
            this.Controls.Add(this.btnCantidad);
            this.Controls.Add(this.label10);
            this.Font = new System.Drawing.Font("Lucida Sans Unicode", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "cajas";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Cantidad";
            this.Load += new System.EventHandler(this.cajas_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnCierra;
        private System.Windows.Forms.TextBox txtCantidad;
        private System.Windows.Forms.Button btnCantidad;
        private System.Windows.Forms.Label label10;
    }
}