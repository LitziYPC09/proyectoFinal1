using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using Xceed.Words.NET; // Necesitas agregar la librería DocX
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using System.Data.SqlClient;

namespace ProyectoFinal_I
{
    public partial class Form1 : Form
    {
        private TextBox promptTextBox;
        private TextBox resultadoTextBox;
        private Button consultarButton;
        private Button guardarButton;
        private Button guardarWordButton;
        private DataGridView historialDataGridView;
        private readonly HttpClient _httpClient;
        private string connectionString = "Server=DESKTOP-MC98NOL\\SQLEXPRESS;Database=investigacionesdb;Trusted_Connection=True;";

        public Form1()
        {

            InitializeComponent();
            _httpClient = new HttpClient();
        }


        private void InitializeComponent()
        {
            this.promptTextBox = new TextBox();
            this.resultadoTextBox = new TextBox();
            this.consultarButton = new Button();
            this.guardarButton = new Button();
            this.historialDataGridView = new DataGridView();
            this.SuspendLayout();

            // promptTextBox
            this.promptTextBox.Location = new System.Drawing.Point(20, 20);
            this.promptTextBox.Multiline = true;
            this.promptTextBox.Size = new System.Drawing.Size(400, 60);
            // this.promptTextBox.PlaceholderText = "Ingresa el prompt aquí...";

            // consultarButton
            this.consultarButton.Location = new System.Drawing.Point(430, 20);
            this.consultarButton.Size = new System.Drawing.Size(100, 30);
            this.consultarButton.Text = "Consultar";
            this.consultarButton.Click += new EventHandler(this.ConsultarYGuardar_Click);
            

            // resultadoTextBox
            this.resultadoTextBox.Location = new System.Drawing.Point(20, 100);
            this.resultadoTextBox.Multiline = true;
            this.resultadoTextBox.Size = new System.Drawing.Size(510, 150);
            this.resultadoTextBox.ReadOnly = true;

            // guardarButton
            this.guardarButton.Location = new System.Drawing.Point(430, 260);
            this.guardarButton.Size = new System.Drawing.Size(100, 30);
            this.guardarButton.Text = "Guardar";


            // this.guardarButton.Click += new EventHandler(this.GuardarButton_Click);
            this.guardarWordButton = new Button();
            this.guardarWordButton.Location = new System.Drawing.Point(360, 510);
            this.guardarWordButton.Size = new System.Drawing.Size(100, 30);
            this.guardarWordButton.Text = "Guardar WORD";
            this.guardarWordButton.Click += new EventHandler(this.GuardarWordButton_Click);
            this.Controls.Add(this.guardarWordButton);

            this.guardarWordButton = new Button();
            this.guardarWordButton.Location = new System.Drawing.Point(150, 510);
            this.guardarWordButton.Size = new System.Drawing.Size(100, 30);
            this.guardarWordButton.Text = "Guardar PPT";
            this.guardarWordButton.Click += new EventHandler(this.GuardarPowerPointButton_Click);
            this.Controls.Add(this.guardarWordButton);

            // historialDataGridView
            this.historialDataGridView.Location = new System.Drawing.Point(20, 310);
            this.historialDataGridView.Size = new System.Drawing.Size(510, 150);
            this.historialDataGridView.ReadOnly = true;

            // InterfazInvestigacion
            this.ClientSize = new System.Drawing.Size(650, 580);
            this.Controls.Add(this.promptTextBox);
            this.Controls.Add(this.consultarButton);
            this.Controls.Add(this.resultadoTextBox);
            this.Controls.Add(this.guardarButton);
            this.Controls.Add(this.historialDataGridView);
            this.Text = "Investigación de Temas";
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private async void ConsultarButton_Click(object sender, EventArgs e)
        {
            string prompt = promptTextBox.Text;
            if (string.IsNullOrWhiteSpace(prompt))
            {
                MessageBox.Show("El prompt no puede estar vacío.");
                return;
            }

            string resultado = await ConsultarGroqAsync(prompt);
            resultadoTextBox.Text = resultado;


        }

        public async Task<string> ConsultarGroqAsync(string prompt)
        {
            try
            {
                string apiKey = "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";
                string endpoint = "https://api.groq.com/openai/v1/chat/completions";

                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Authorization =
                    new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey);

                    var requestBody = new
                    {
                        model = "llama3-70b-8192",
                        messages = new[]
                        {
                            new Dictionary<string, string>
                            {
                                { "role", "user" },
                                { "content", prompt }
                            }
                        },
                        temperature = 0.7
                    };

                    var json = JsonConvert.SerializeObject(requestBody);
                    var content = new StringContent(json, Encoding.UTF8, "application/json");

                    var response = await client.PostAsync(endpoint, content);

                    if ((int)response.StatusCode == 429)
                        return "⚠️ Demasiadas solicitudes. Intenta de nuevo más tarde.";

                    if ((int)response.StatusCode == 401)
                        return "❌ API Key inválida. Verifica tu clave.";

                    if ((int)response.StatusCode == 402)
                        return "❌ Error 402: No tienes créditos o acceso al modelo.";

                    if ((int)response.StatusCode == 400)
                    {
                        string errorContent = await response.Content.ReadAsStringAsync();
                        return $"❌ Error 400: Solicitud inválida. Detalles: {errorContent}";
                    }

                    response.EnsureSuccessStatusCode();

                    string jsonResult = await response.Content.ReadAsStringAsync();
                    dynamic result = JsonConvert.DeserializeObject(jsonResult);

                    return result.choices[0].message.content.ToString().Trim();
                }
            }
            catch (HttpRequestException ex)
            {
                return $"❌ Error HTTP: {ex.Message}";
            }
            catch (Exception ex)
            {
                return $"❌ Error general: {ex.Message}";
            }
        }


        private void GuardarWordButton_Click(object sender, EventArgs e)
        {
            string prompt = promptTextBox.Text;
            string resultado = resultadoTextBox.Text;

            if (string.IsNullOrWhiteSpace(prompt) || string.IsNullOrWhiteSpace(resultado))
            {
                MessageBox.Show("El prompt y el resultado no pueden estar vacíos.");
                return;
            }

            GuardarEnWord(prompt, resultado);
        }

        private void GuardarEnWord(string prompt, string resultado)
        {
            try
            {

                string rutaDestino = "C:\\Users\\yasmi\\OneDrive\\Documentos\\Documentos Guatados por IA";
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string fileName = $"resultado_investigacion_{timestamp}.docx";
                string filePath = Path.Combine(rutaDestino, fileName);

                using (var doc = DocX.Create(filePath))
                {
                    doc.InsertParagraph("Prompt:").Bold();
                    doc.InsertParagraph(prompt);
                    doc.InsertParagraph("\nResultado:").Bold();
                    doc.InsertParagraph(resultado);
                    doc.Save();
                }

                MessageBox.Show($"Documento guardado automáticamente en:\n{filePath}", "Guardado exitoso");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Error al guardar el archivo: {ex.Message}");
            }
        }



        public void GuardarEnPowerPoint(string prompt, string resultado, string rutaDestino)
        {
            try
            {
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string fileName = $"presentacion_investigacion_{timestamp}.pptx";
                string filePath = Path.Combine(rutaDestino, fileName);

                using (PresentationDocument presentationDoc = PresentationDocument.Create(filePath, PresentationDocumentType.Presentation))
                {
                    PresentationPart presentationPart = presentationDoc.AddPresentationPart();
                    presentationPart.Presentation = new Presentation();

                    SlidePart slide1 = CrearDiapositiva(presentationPart, "Prompt", prompt);
                    SlidePart slide2 = CrearDiapositiva(presentationPart, "Resultado", resultado);

                    presentationPart.Presentation.SlideIdList = new SlideIdList();

                    uint slideId = 256;

                    foreach (var slide in new[] { slide1, slide2 })
                    {
                        SlideId slideIdElement = new SlideId()
                        {
                            Id = slideId++,
                            RelationshipId = presentationPart.GetIdOfPart(slide)
                        };
                        presentationPart.Presentation.SlideIdList.Append(slideIdElement);
                    }

                    presentationPart.Presentation.Save();
                }

                MessageBox.Show($"Presentación guardada en:\n{filePath}", "Guardado exitoso");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Error al guardar el PowerPoint: {ex.Message}");
            }
        }

        private SlidePart CrearDiapositiva(PresentationPart presentationPart, string titulo, string contenido)
        {
            SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();
            slidePart.Slide = new Slide(
                new CommonSlideData(
                    new ShapeTree(
                        new NonVisualGroupShapeProperties(
                            new NonVisualDrawingProperties() { Id = 1, Name = "" },
                            new NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                        new GroupShapeProperties(new A.TransformGroup()),
                        CrearTexto(titulo, 0, 0, 720, 100),
                        CrearTexto(contenido, 0, 100, 720, 500)
                    )
                )
            );
            return slidePart;
        }

        private Shape CrearTexto(string texto, int x, int y, int cx, int cy)
        {
            return new Shape(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties() { Id = 2, Name = "Texto" },
                    new NonVisualShapeDrawingProperties(new A.ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties()),
                new ShapeProperties(
                    new A.Transform2D(
                        new A.Offset() { X = x * 9525, Y = y * 9525 },
                        new A.Extents() { Cx = cx * 9525, Cy = cy * 9525 })),
                new TextBody(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(new A.Run(new A.Text(texto)))
                )
            );
        }

        private void GuardarPowerPointButton_Click(object sender, EventArgs e)
        {
            string prompt = promptTextBox.Text;
            string resultado = resultadoTextBox.Text;

            if (string.IsNullOrWhiteSpace(prompt) || string.IsNullOrWhiteSpace(resultado))
            {
                MessageBox.Show("El prompt y el resultado no pueden estar vacíos.");
                return;
            }

            // Usar carpeta predefinida
            string rutaDestino = "C:\\Users\\yasmi\\OneDrive\\Documentos\\Documentos Guatados por IA";

            // Crear carpeta si no existe
            if (!Directory.Exists(rutaDestino))
            {
                Directory.CreateDirectory(rutaDestino);
            }

            // Llamar al método para guardar en PowerPoint
            GuardarEnPowerPoint(prompt, resultado, rutaDestino);
        }

        private void GuardarEnBaseDeDatos(string prompt, string resultado)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    string query = "INSERT INTO Investigaciones (Prompt, Resultado) VALUES (@Prompt, @Resultado)";
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@Prompt", prompt);
                        command.Parameters.AddWithValue("@Resultado", resultado);
                        command.ExecuteNonQuery();
                    }
                }

                MessageBox.Show("✅ Guardado en la base de datos con éxito.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Error al guardar en la base de datos: {ex.Message}");
            }
        }

        private void GuardarDBButton_Click(object sender, EventArgs e)
        {
            string prompt = promptTextBox.Text;
            string resultado = resultadoTextBox.Text;

            if (string.IsNullOrWhiteSpace(prompt) || string.IsNullOrWhiteSpace(resultado))
            {
                MessageBox.Show("El prompt y el resultado no pueden estar vacíos.");
                return;
            }

            GuardarEnBaseDeDatos(prompt, resultado);
           
        }

        private void ConsultarYGuardar_Click(object sender, EventArgs e)
        {
            ConsultarButton_Click(sender, e);
            GuardarDBButton_Click(sender, e);
        }




    }
}