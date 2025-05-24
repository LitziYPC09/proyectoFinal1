using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using Xceed.Words.NET; 
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using Xceed.Document.NET;

namespace ProyectoFinal_I
{
    public partial class Form1 : Form
    {
        private TextBox promptTextBox;
        private TextBox resultadoTextBox;
        private Button consultarButton;
        private Button guardarButton;
        private Button guardarWordButton;
        private Button guardarPowerPointButton;
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
            this.guardarWordButton = new Button();
            this.guardarPowerPointButton = new Button();
            this.historialDataGridView = new DataGridView();

            this.SuspendLayout();

            // promptTextBox
            this.promptTextBox.Location = new System.Drawing.Point(20, 20);
            this.promptTextBox.Multiline = true;
            this.promptTextBox.Size = new System.Drawing.Size(400, 60);
            this.BackColor = Color.Black;

            // consultarButton
            this.consultarButton.Location = new System.Drawing.Point(430, 20);
            this.consultarButton.Size = new System.Drawing.Size(100, 30);
            this.consultarButton.Text = "Consultar";
            this.consultarButton.BackColor = Color.DarkBlue;
            this.consultarButton.ForeColor = Color.White;
            this.consultarButton.FlatStyle = FlatStyle.Flat;
            this.consultarButton.Click += new EventHandler(this.ConsultarYGuardar_Click);

            // resultadoTextBox
            this.resultadoTextBox.Location = new System.Drawing.Point(20, 100);
            this.resultadoTextBox.Multiline = true;
            this.resultadoTextBox.Size = new System.Drawing.Size(510, 150);
            this.resultadoTextBox.ReadOnly = true;
            this.resultadoTextBox.BackColor = Color.DarkGray;

            // guardarButton
            this.guardarButton.Location = new System.Drawing.Point(430, 260);
            this.guardarButton.Size = new System.Drawing.Size(100, 30);
            this.guardarButton.Text = "Guardar";
            this.guardarButton.BackColor = Color.DarkBlue;
            this.guardarButton.ForeColor = Color.White;
            this.guardarButton.FlatStyle = FlatStyle.Flat;

            // guardarWordButton
            this.guardarWordButton.Location = new System.Drawing.Point(360, 510);
            this.guardarWordButton.Size = new System.Drawing.Size(100, 30);
            this.guardarWordButton.Text = "Guardar WORD";
            this.guardarWordButton.BackColor = Color.DarkBlue;
            this.guardarWordButton.ForeColor = Color.White;
            this.guardarWordButton.FlatStyle = FlatStyle.Flat;
            this.guardarWordButton.Click += new EventHandler(this.GuardarWordButton_Click);
            this.Controls.Add(this.guardarWordButton);

            // guardarPowerPointButton
            this.guardarPowerPointButton.Location = new System.Drawing.Point(150, 510);
            this.guardarPowerPointButton.Size = new System.Drawing.Size(100, 30);
            this.guardarPowerPointButton.Text = "Guardar PPT";
            this.guardarPowerPointButton.BackColor = Color.DarkBlue;
            this.guardarPowerPointButton.ForeColor = Color.White;
            this.guardarPowerPointButton.FlatStyle = FlatStyle.Flat;
            this.guardarPowerPointButton.Click += new EventHandler(this.GuardarPowerPointButton_Click);
            this.Controls.Add(this.guardarPowerPointButton);

            // historialDataGridView
            this.historialDataGridView.Location = new System.Drawing.Point(20, 310);
            this.historialDataGridView.Size = new System.Drawing.Size(510, 150);
            this.historialDataGridView.ReadOnly = true;
            this.historialDataGridView.BackColor = Color.DarkGray;

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
                        return "⚠ Demasiadas solicitudes. Intenta de nuevo más tarde.";

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


        private  async void GuardarWordButton_Click(object sender, EventArgs e)
        {
            string prompt = promptTextBox.Text;
            string resultado = resultadoTextBox.Text;

            if (string.IsNullOrWhiteSpace(prompt) || string.IsNullOrWhiteSpace(resultado))
            {
                MessageBox.Show("Por favor, completa el prompt y espera el resultado antes de guardar.", "Advertencia");
                return;
            }

            try
            {
                string rutaDestino = @"C:\Users\yasmi\OneDrive\Documentos\Documentos Guatados por IA";
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string fileName = $"resultado_investigacion_{timestamp}.docx";
                string filePath = Path.Combine(rutaDestino, fileName);

                CrearPlantillaWord(filePath, prompt, resultado);

                MessageBox.Show($"Documento guardado exitosamente en:\n{filePath}", "Guardado exitoso");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Error al guardar el archivo: {ex.Message}");
            }
        }


        //private void GuardarWordButton(string prompt, string resultado)
        //{
        //    try
        //    {
        //        string rutaDestino = @"C:\Users\yasmi\OneDrive\Documentos\Documentos Guatados por IA";
        //        string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        //        string fileName = $"resultado_investigacion_{timestamp}.docx";
        //        string filePath = Path.Combine(rutaDestino, fileName);

        //        // Llamamos a la función de plantilla sin colores
        //        CrearPlantillaWord(filePath, prompt, resultado);

        //        MessageBox.Show($"Documento guardado exitosamente en:\n{filePath}", "Guardado exitoso");
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show($"❌ Error al guardar el archivo: {ex.Message}");
        //    }
        //}


        private void CrearPlantillaWord(string filePath, string prompt, string resultado)
        {
            using (var doc = DocX.Create(filePath))
            {
                // Título principal
                var titulo = doc.InsertParagraph("Reporte de Investigación Generado por IA")
                                .FontSize(18)
                                .Bold();
                titulo.Alignment = Alignment.center;

                doc.InsertParagraph(Environment.NewLine);

                // Sección Prompt
                doc.InsertParagraph("Prompt:")
                   .Bold()
                   .FontSize(14);
                doc.InsertParagraph(prompt)
                   .FontSize(12)
                   .SpacingAfter(10);

                // Separador
                var separador = doc.InsertParagraph(new string('-', 50));
                separador.Alignment = Alignment.center;

                // Sección Resultado
                doc.InsertParagraph("Resultado:")
                   .Bold()
                   .FontSize(14);
                doc.InsertParagraph(resultado)
                   .FontSize(12)
                   .SpacingAfter(10);

                doc.Save();
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

                    List<SlidePart> slides = new List<SlidePart>();

                    // Primera diapositiva con título "Prompt"
                    slides.Add(CrearDiapositiva(presentationPart, "Prompt", prompt));

                    // Dividir el contenido en varias partes
                    var partesContenido = DividirContenidoEnPartes(resultado, 80); // 80 palabras por diapositiva

                    for (int i = 0; i < partesContenido.Count; i++)
                    {
                        string titulo = $"Resultado (Parte {i + 1}";
                        slides.Add(CrearDiapositiva(presentationPart, titulo, partesContenido[i]));
                    }

                    // Crear SlideIdList
                    presentationPart.Presentation.SlideIdList = new SlideIdList();
                    uint slideId = 256;

                    foreach (var slide in slides)
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

        private List<string> DividirContenidoEnPartes(string contenido, int maxPalabrasPorParte)
        {
            List<string> partes = new List<string>();
            var palabras = contenido.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            for (int i = 0; i < palabras.Length; i += maxPalabrasPorParte)
            {
                var parte = string.Join(" ", palabras.Skip(i).Take(maxPalabrasPorParte));
                partes.Add(parte);
            }

            return partes;
        }

        private SlidePart CrearDiapositiva(PresentationPart presentationPart, string titulo, string contenido)
        {
            SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();

            // Crear el fondo (color claro)
            var background = new Background(
                new BackgroundProperties(
                    new A.SolidFill(new A.RgbColorModelHex() { Val = ObtenerHexDesdeNombre("LightGray") })
                )
            );

            // Crear las formas de texto: título y contenido
            var tituloShape = CrearTexto(titulo, 170, 40, 620, 80, true);
            var contenidoShape = CrearTexto(contenido, 50, 160, 620, 400, false);
            //var lineaDecorativa = CrearLineaDecorativa(50, 130, 620, 4, "SteelBlue");

            var shapeTree = new ShapeTree(
                new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties() { Id = 0, Name = "Slide Root" },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new GroupShapeProperties(new A.TransformGroup()),
                tituloShape,
               // lineaDecorativa,
                contenidoShape
            );

            // Crear CommonSlideData e insertar fondo como primer elemento
            var commonSlideData = new CommonSlideData(shapeTree);
            commonSlideData.InsertAt(background, 0);

            // Crear Slide con ColorMapOverride (para buen contraste)
            slidePart.Slide = new Slide(commonSlideData)
            {
                ColorMapOverride = new ColorMapOverride(new A.MasterColorMapping())
            };

            slidePart.Slide.Save();
            return slidePart;
        }

        private Shape CrearTexto(string texto, int x, int y, int cx, int cy, bool esTitulo = false)
        {
            uint shapeId = (uint)(esTitulo ? 1 : 2);
            int fontSize = esTitulo ? 3200 : 1800;
            string colorNombre = esTitulo ? "DarkBlue" : "Black";

            var runProperties = new A.RunProperties()
            {
                FontSize = fontSize,
                Bold = esTitulo
            };
            runProperties.Append(new A.SolidFill(new A.RgbColorModelHex() { Val = ObtenerHexDesdeNombre(colorNombre) }));

            var paragraphProps = new A.ParagraphProperties()
            {
                Alignment = esTitulo ? A.TextAlignmentTypeValues.Center : A.TextAlignmentTypeValues.Left
            };

            // Asegúrate que el texto no sea null o vacío
            if (string.IsNullOrEmpty(texto))
            {
                texto = "(Sin contenido)";
            }

            return new Shape(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties() { Id = shapeId, Name = esTitulo ? "Título" : "Contenido" },
                    new NonVisualShapeDrawingProperties(new A.ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties()),
                new ShapeProperties(
                    new A.Transform2D(
                        new A.Offset() { X = x * 9525, Y = y * 9525 },
                        new A.Extents() { Cx = cx * 9525, Cy = cy * 9525 })),
                new TextBody(
                    new A.BodyProperties()
                    {
                        Anchor = esTitulo ? A.TextAnchoringTypeValues.Center : A.TextAnchoringTypeValues.Top,
                        // Wrap = A.TextWrappingValues.Square (opcional)
                    },
                    new A.ListStyle(),
                    new A.Paragraph(paragraphProps, new A.Run(runProperties, new A.Text(texto)))
                )
            );
        }


        private string ObtenerHexDesdeNombre(string nombreColor)
        {
            var color = System.Drawing.Color.FromName(nombreColor);
            return $"{color.R:X2}{color.G:X2}{color.B:X2}";
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