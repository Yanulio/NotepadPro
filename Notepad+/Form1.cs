using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Notepad_
{
    /// <summary>
    ///  Класс для основной работы с формой.
    /// </summary>

    public partial class Notepad : Form
    {

        /// <summary>
        /// Конструктор формы, в котором форма принимает нужные настройки цвета и таймера для автосохранения.
        /// </summary>

        public Notepad()
        {
            InitializeComponent();
            ChangingAllColors();
            autoSaveTimer.Interval = Properties.Settings.Default.autoSaveNum;
            autoSaveTimer.Enabled = Properties.Settings.Default.autoSaveOn;
        }

        /// <summary>
        /// Метод события загрузки формы, в котором происходит загрузка ранее открытых вкладок.
        /// </summary>
        /// <param name="sender">Объект, над которым совершено событие.</param>
        /// <param name="e">Информация о событии.</param>

        private void Notepad_Load(object sender, EventArgs e)
        {
            if (Application.OpenForms.Count == 1)
            {
                for (int i = 0; i < Properties.Settings.Default.files.Count; i++)
                {
                    NewTab();
                    tabControl.SelectedTab.Name = Properties.Settings.Default.files[i];
                    try
                    {
                        if (File.Exists(tabControl.SelectedTab.Name))
                        {
                            Control control = tabControl.Controls[tabControl.SelectedIndex].Controls[0];
                            RichTextBox richTextBox = control as RichTextBox;
                            richTextBox.TextChanged += new EventHandler(TextChangedHandler);
                            if (Path.GetExtension(tabControl.SelectedTab.Name) == ".rtf")
                                richTextBox.LoadFile(tabControl.SelectedTab.Name, RichTextBoxStreamType.RichText);
                            else
                            {
                                StreamReader sr = new StreamReader(tabControl.SelectedTab.Name);
                                richTextBox.Text = sr.ReadToEnd();
                                sr.Close();
                            }
                        }
                        else
                        {
                            MessageBox.Show($"Файла {tabControl.SelectedTab.Name} больше не существует, загрузить файл невозможно.", "Внимание!");
                            RemoveTab(tabControl.SelectedTab);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Произошла ошибка! {ex.Message}", "Внимание!");
                        RemoveTab(tabControl.SelectedTab);
                    }
                }
                if (tabControl.TabPages.Count != 0)
                    tabControl.Visible = true;
            }
        }

        /// <summary>
        /// Метод, удаляющий одну из вкладок.
        /// </summary>
        /// <param name="tabPage">Вкладка для удаления.</param>

        private void RemoveTab(TabPage tabPage)
        {
            try
            {
                tabControl.TabPages.Remove(tabPage);
                string realName = tabPage.Name;
                if (tabPage.Name[tabPage.Name.Length - 1] == '*')
                    realName = tabPage.Name.Substring(0, tabPage.Name.Length - 1);
                Properties.Settings.Default.files.Remove(realName);
                Properties.Settings.Default.Save();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка! {ex.Message}", "Внимание!");
            }
        }

        /// <summary>
        /// Метод, устанавливающий несколько настроек для rich text box.
        /// </summary>
        /// <param name="richTextBox">Текст бокс, у которого меняются настройки.</param>

        private void richTextBoxSettings(RichTextBox richTextBox)
        {
            richTextBox.Dock = DockStyle.Fill;
            richTextBox.BackColor = Properties.Settings.Default.richTextBoxColor;
            richTextBox.ContextMenuStrip = contextMenuStrip;
        }

        /// <summary>
        /// Метод, создающий новую вкладку в tab control.
        /// </summary>

        private void NewTab()
        {
            try
            {
                TabPage newTabPage = new TabPage();
                newTabPage.Text = "Новая вкладка (" + (tabControl.TabPages.Count + 1) + ")";
                tabControl.TabPages.Add(newTabPage);
                RichTextBox richTextBox = new RichTextBox();
                newTabPage.Controls.Add(richTextBox);
                tabControl.SelectedTab = tabControl.TabPages[tabControl.TabPages.Count - 1];
                tabControl.SelectedTab.BorderStyle = BorderStyle.FixedSingle;
                richTextBoxSettings(richTextBox);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка! {ex.Message}", "Внимание!");
            }
        }

        /// <summary>
        /// Метод события, открывающий выбранный пользователем файл в новой вкладке.
        /// </summary>
        /// <param name="sender">Объект, над которым совершено событие.</param>
        /// <param name="e">Информация о событии.</param>

        private void ВНовойВкладкеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog()
            {
                InitialDirectory = Directory.GetCurrentDirectory()
            };
            if (dialog.ShowDialog() != DialogResult.OK)
                return;
            if (tabControl.TabPages.Count == 0)
                tabControl.Visible = true;
            NewTab();
            TabPage currentTab = tabControl.SelectedTab;
            try
            {
                Control control = tabControl.Controls[tabControl.SelectedIndex].Controls[0];
                RichTextBox richTextBox = control as RichTextBox;
                string fileName = dialog.FileName;
                string extension = Path.GetExtension(fileName);
                if (extension == ".rtf")
                    richTextBox.LoadFile(fileName, RichTextBoxStreamType.RichText);
                else
                {
                    StreamReader sr = new StreamReader(fileName);
                    richTextBox.Text = sr.ReadToEnd();
                    sr.Close();
                }
                currentTab.Name = fileName;
                richTextBox.TextChanged += new EventHandler(TextChangedHandler);
                Properties.Settings.Default.files.Add(fileName);
                if (Properties.Settings.Default.files[0] == " ")
                    Properties.Settings.Default.files.RemoveAt(0);
                Properties.Settings.Default.Save();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка! {ex.Message}", "Внимание!");
                tabControl.TabPages.Remove(tabControl.SelectedTab);
                tabControl.Visible = false;
            }
        }

        /// <summary>
        /// Метод события, открывающий выбранный пользователем файл в новом окне.
        /// </summary>
        /// <param name="sender">Объект, над которым совершено событие.</param>
        /// <param name="e">Информация о событии.</param>

        private void ВНовомОкнеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog() { InitialDirectory = Directory.GetCurrentDirectory() };
            if (dialog.ShowDialog() != DialogResult.OK)
                return;
            Properties.Settings.Default.Save();
            Notepad newForm = new Notepad();
            TabPage newTabPage = new TabPage();
            newTabPage.Text = "Новая вкладка (" + "1" + ")";
            newForm.tabControl.TabPages.Add(newTabPage);
            RichTextBox richTextBox = new RichTextBox();
            newTabPage.Controls.Add(richTextBox);
            newForm.tabControl.SelectedTab = newForm.tabControl.TabPages[newForm.tabControl.TabPages.Count - 1];
            newForm.tabControl.SelectedTab.BorderStyle = BorderStyle.FixedSingle;
            newForm.tabControl.Visible = true;
            richTextBoxSettings(richTextBox);
            try
            {
                if (Path.GetExtension(dialog.FileName) == ".rtf")
                    richTextBox.LoadFile(dialog.FileName, RichTextBoxStreamType.RichText);
                else
                {
                    StreamReader sr = new StreamReader(dialog.FileName);
                    richTextBox.Text = sr.ReadToEnd();
                    sr.Close();
                }
                newForm.tabControl.SelectedTab.Name = dialog.FileName;
                richTextBox.TextChanged += new EventHandler(TextChangedHandler);
                Properties.Settings.Default.files.Add(dialog.FileName);
                if (Properties.Settings.Default.files[0] == " ")
                    Properties.Settings.Default.files.RemoveAt(0);
                Properties.Settings.Default.Save();
                newForm.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка! {ex.Message}", "Внимание!");
                tabControl.TabPages.Remove(tabControl.SelectedTab);
                tabControl.Visible = false;
            }
        }

        /// <summary>
        /// Метод события, закрывающий выбранную на данный момент вкладку.
        /// </summary>
        /// <param name="sender">Объект, над которым совершено событие.</param>
        /// <param name="e">Информация о событии.</param>

        private void ЗакрытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (tabControl.TabPages.Count != 0)
                {
                    if (tabControl.SelectedTab.Name[tabControl.SelectedTab.Name.Length - 1] == '*')
                        Properties.Settings.Default.files.Remove(tabControl.SelectedTab.Name.Substring(0, tabControl.SelectedTab.Name.Length - 1));
                    else
                        Properties.Settings.Default.files.Remove(tabControl.SelectedTab.Name);
                    Properties.Settings.Default.Save();
                    tabControl.TabPages.Remove(tabControl.SelectedTab);
                    if (tabControl.TabPages.Count == 0)
                        tabControl.Visible = false;
                }
                else
                    MessageBox.Show("Нет активных вкладок для закрытия.", "Внимание!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка! {ex.Message}", "Внимание!");
            }
        }

        /// <summary>
        /// Метод события, закрывающий все открытые в окне вкладкию.
        /// </summary>
        /// <param name="sender">Объект, над которым совершено событие.</param>
        /// <param name="e">Информация о событии.</param>
        private void ЗакрытьВсёToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                int count = tabControl.TabPages.Count;
                if (count != 0)
                {
                    for (int i = 0; i < count; i++)
                    {
                        if (tabControl.SelectedTab.Name[tabControl.TabPages[0].Name.Length - 1] == '*')
                            Properties.Settings.Default.files.Remove(tabControl.TabPages[0].Name.Substring(0, tabControl.TabPages[0].Name.Length - 1));
                        else
                            Properties.Settings.Default.files.Remove(tabControl.TabPages[0].Name);
                        Properties.Settings.Default.Save();
                        tabControl.TabPages.Remove(tabControl.TabPages[0]);
                    }
                    tabControl.Visible = false;
                }
                else
                    MessageBox.Show("Нет активных вкладок для закрытия.", "Внимание!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка! {ex.Message}", "Внимание!");
            }
        }

        /// <summary>
        /// Метод события, отвечающий за выход из приложения.
        /// </summary>
        /// <param name="sender">Объект, над которым совершено событие.</param>
        /// <param name="e">Информация о событии.</param>

        private void ВыходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        /// <summary>
        /// Метод события, выделяющий во всей вкладке текст.
        /// </summary>
        /// <param name="sender">Объект, над которым совершено событие.</param>
        /// <param name="e">Информация о событии.</param>

        private void ВыбратьВесьТекстToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                Control control = tabControl.Controls[tabControl.SelectedIndex].Controls[0];
                RichTextBox richTextBox = control as RichTextBox;
                richTextBox.SelectAll();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка! {ex.Message}", "Внимание!");
            }
        }

        /// <summary>
        /// Метод события, удаляюшщий выделенный в активной вкладке текст.
        /// </summary>
        /// <param name="sender">Объект, над которым совершено событие.</param>
        /// <param name="e">Информация о событии.</param>

        private void ВырезатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                Control control = tabControl.Controls[tabControl.SelectedIndex].Controls[0];
                RichTextBox richTextBox = control as RichTextBox;
                richTextBox.SelectedText = "";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка! {ex.Message}", "Внимание!");
            }
        }

        /// <summary>
        /// Метод события, копирующий выделенный в активной вкладке текст в буфер обмена.
        /// </summary>
        /// <param name="sender">Объект, над которым совершено событие.</param>
        /// <param name="e">Информация о событии.</param>

        private void КопироватьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                Control control = tabControl.Controls[tabControl.SelectedIndex].Controls[0];
                RichTextBox richTextBox = control as RichTextBox;
                Clipboard.SetText(richTextBox.SelectedText);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка! {ex.Message}", "Внимание!");
            }
        }

        /// <summary>
        /// Метод события, вставляющий в активную вкладку текст из буфера обмена.
        /// </summary>
        /// <param name="sender">Объект, над которым совершено событие.</param>
        /// <param name="e">Информация о событии.</param>

        private void ВставитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                Control control = tabControl.Controls[tabControl.SelectedIndex].Controls[0];
                RichTextBox richTextBox = control as RichTextBox;
                richTextBox.Paste();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка! {ex.Message}", "Внимание!");
            }
        }

        /// <summary>
        /// Метод события, сохраняющий текст вкладки в файл, который был открыт в активной вкладке. 
        /// Если файла не существует, об этом сообщается пользователю.
        /// </summary>
        /// <param name="sender">Объект, над которым совершено событие.</param>
        /// <param name="e">Информация о событии.</param>

        private void СохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (tabControl.TabPages.Count == 0)
                MessageBox.Show("Нет активных вкладок для сохранения.", "Внимание!");
            else
            {
                try
                {
                    string realName = tabControl.SelectedTab.Name;
                    if (tabControl.SelectedTab.Name[tabControl.SelectedTab.Name.Length - 1] == '*')
                        realName = tabControl.SelectedTab.Name.Substring(0, tabControl.SelectedTab.Name.Length - 1);
                    if (File.Exists(realName))
                    {
                        Control control = tabControl.Controls[tabControl.SelectedIndex].Controls[0];
                        RichTextBox richTextBox = control as RichTextBox;
                        if (Path.GetExtension(realName) == ".rtf")
                            richTextBox.SaveFile(realName);
                        else
                        {
                            StreamWriter sw = new StreamWriter(realName);
                            sw.Write(richTextBox.Text);
                            sw.Close();
                        }
                        tabControl.SelectedTab.Name = realName;
                    }
                    else
                        MessageBox.Show($"Файла {realName} больше не существует, попробуйте " +
                            $"воспользоваться функцией \"Сохранить как\".", "Внимание!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Произошла ошибка! {ex.Message}", "Внимание!");
                }
            }
        }

        /// <summary>
        /// Метод события, сохраняющий текст из всех активных вкладок в соответствующие им файлы.
        /// </summary>
        /// <param name="sender">Объект, над которым совершено событие.</param>
        /// <param name="e">Информация о событии.</param>

        private void СохранитьВсёtoolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (tabControl.TabPages.Count == 0)
                MessageBox.Show("Нет активных вкладок для сохранения.", "Внимание!");
            else
            {
                for (int i = 0; i < tabControl.TabPages.Count; i++)
                {
                    try
                    {
                        string realName = tabControl.TabPages[i].Name;
                        if (tabControl.TabPages[i].Name[tabControl.TabPages[i].Name.Length - 1] == '*')
                            realName = tabControl.TabPages[i].Name.Substring(0, tabControl.TabPages[i].Name.Length - 1);
                        if (File.Exists(realName))
                        {
                            Control control = tabControl.Controls[i].Controls[0];
                            RichTextBox richTextBox = control as RichTextBox;
                            if (Path.GetExtension(realName) == ".rtf")
                                richTextBox.SaveFile(realName);
                            else
                            {
                                StreamWriter sw = new StreamWriter(realName);
                                sw.Write(richTextBox.Text);
                                sw.Close();
                            }
                            tabControl.TabPages[i].Name = realName;
                        }
                        else
                            MessageBox.Show($"Файла {realName} больше не существует, попробуйте " +
                                $"воспользоваться функцией \"Сохранить как\".", "Внимание!");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Произошла ошибка! {ex.Message}", "Внимание!");
                    }
                }
            }
        }

        /// <summary>
        /// Метод события, позволяющий пользователю выбрать, куда созранять текст из вкладки.
        /// </summary>
        /// <param name="sender">Объект, над которым совершено событие.</param>
        /// <param name="e">Информация о событии.</param>

        private void СохранитьКакToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (tabControl.TabPages.Count == 0)
                MessageBox.Show("Нет активных вкладок для сохранения.", "Внимание!");
            else
            {
                try
                {
                    Control control = tabControl.Controls[tabControl.SelectedIndex].Controls[0];
                    RichTextBox richTextBox = control as RichTextBox;
                    SaveFileDialog saveFileDialog = new SaveFileDialog();
                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        if (Path.GetExtension(saveFileDialog.FileName.ToString()) == ".rtf")
                            richTextBox.SaveFile(saveFileDialog.FileName.ToString());
                        else
                        {
                            StreamWriter sw = new StreamWriter(saveFileDialog.FileName.ToString());
                            sw.Write(richTextBox.Text);
                            sw.Close();
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Произошла ошибка! {ex.Message}", "Внимание!");
                }
            }
        }

        /// <summary>
        /// Метод события, меняющий цветовую схему приложения на более тёмную.
        /// </summary>
        /// <param name="sender">Объект, над которым совершено событие.</param>
        /// <param name="e">Информация о событии.</param>

        private void ТёмнаяToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.menuColor = SystemColors.WindowFrame;
            Properties.Settings.Default.menuFontColor = SystemColors.ControlLight;
            Properties.Settings.Default.richTextBoxColor = Color.Silver;
            Properties.Settings.Default.richTextBoxFontColor = SystemColors.WindowText;
            Properties.Settings.Default.contextMenuColor = Color.Silver;
            Properties.Settings.Default.backgroundColor = Color.Silver;
            Properties.Settings.Default.Save();
            ChangingAllColors();
        }

        /// <summary>
        /// Метод события, меняющий цветовую схему приложения на более светлую.
        /// </summary>
        /// <param name="sender">Объект, над которым совершено событие.</param>
        /// <param name="e">Информация о событии.</param>

        private void СветлаяToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.richTextBoxColor = SystemColors.Window;
            Properties.Settings.Default.richTextBoxFontColor = SystemColors.WindowText;
            Properties.Settings.Default.menuColor = SystemColors.GradientInactiveCaption;
            Properties.Settings.Default.contextMenuColor = SystemColors.ButtonHighlight;
            Properties.Settings.Default.menuFontColor = SystemColors.ControlText;
            Properties.Settings.Default.backgroundColor = SystemColors.Window;
            Properties.Settings.Default.Save();
            ChangingAllColors();
        }

        /// <summary>
        /// Метод события, меняющий цветовую схему приложения на бирюзовую.
        /// </summary>
        /// <param name="sender">Объект, над которым совершено событие.</param>
        /// <param name="e">Информация о событии.</param>

        private void БирюзовыйФлексToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.menuColor = Color.DarkCyan;
            Properties.Settings.Default.menuFontColor = SystemColors.ControlText;
            Properties.Settings.Default.richTextBoxColor = Color.PowderBlue;
            Properties.Settings.Default.richTextBoxFontColor = SystemColors.ControlText;
            Properties.Settings.Default.contextMenuColor = Color.LightCyan;
            Properties.Settings.Default.backgroundColor = Color.PowderBlue;
            Properties.Settings.Default.Save();
            ChangingAllColors();
        }

        /// <summary>
        /// Метод события, меняющий цветовую схему приложения на розовую.
        /// </summary>
        /// <param name="sender">Объект, над которым совершено событие.</param>
        /// <param name="e">Информация о событии.</param>

        private void РозовыйКарнавалToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.menuColor = Color.MediumVioletRed;
            Properties.Settings.Default.menuFontColor = SystemColors.ControlText;
            Properties.Settings.Default.richTextBoxColor = Color.LightPink;
            Properties.Settings.Default.richTextBoxFontColor = SystemColors.ControlText;
            Properties.Settings.Default.contextMenuColor = Color.Pink;
            Properties.Settings.Default.backgroundColor = Color.LightPink;
            Properties.Settings.Default.Save();
            ChangingAllColors();
        }

        /// <summary>
        /// Метод, меняющий все цвета формы на нужные в зависимости от цветовой схемы.
        /// </summary>

        private void ChangingAllColors()
        {
            try
            {
                Menu.BackColor = Properties.Settings.Default.menuColor;
                for (int i = 0; i < Menu.Items.Count; i++)
                    Menu.Items[i].ForeColor = Properties.Settings.Default.menuFontColor;
                AdditionalMenuColors();
                this.BackColor = Properties.Settings.Default.backgroundColor;
                contextMenuStrip.BackColor = Properties.Settings.Default.contextMenuColor;
                int count = tabControl.TabPages.Count;
                for (int i = 0; i < count; i++)
                {
                    Control control = tabControl.Controls[i].Controls[0];
                    RichTextBox richTextBox = control as RichTextBox;
                    richTextBox.BackColor = Properties.Settings.Default.richTextBoxColor;
                    richTextBox.ForeColor = Properties.Settings.Default.richTextBoxFontColor;
                }
                if (tabControl.TabPages.Count == 0)
                    tabControl.Visible = false;
                //Меняю рендер для меню, чтобы подменю тоже окрашивались в нужный цвет (соответсвующие классы ниже).
                if (Properties.Settings.Default.menuColor == SystemColors.GradientInactiveCaption)
                    Menu.Renderer = new GradientInactiveCaptionRenderer();
                if (Properties.Settings.Default.menuColor == SystemColors.WindowFrame)
                    Menu.Renderer = new WindowFrameRenderer();
                if (Properties.Settings.Default.menuColor == Color.DarkCyan)
                    Menu.Renderer = new DarkCyanRenderer();
                if (Properties.Settings.Default.menuColor == Color.MediumVioletRed)
                    Menu.Renderer = new MediumVioletRedRenderer();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка! {ex.Message}", "Внимание!");
            }
        }

        /// <summary>
        /// Метод, меняющий цвет некоторых объектов на нужный цвет (очень извиняюсь за это,
        /// но на тот момент я другого способа не нашла...)
        /// </summary>

        private void AdditionalMenuColors()
        {
            открытьToolStripMenuItem.ForeColor = Properties.Settings.Default.menuFontColor;
            ВНовойВкладкеToolStripMenuItem.ForeColor = Properties.Settings.Default.menuFontColor;
            ВНовомОкнеToolStripMenuItem.ForeColor = Properties.Settings.Default.menuFontColor;
            СохранитьToolStripMenuItem.ForeColor = Properties.Settings.Default.menuFontColor;
            СохранитьКакToolStripMenuItem.ForeColor = Properties.Settings.Default.menuFontColor;
            ЗакрытьToolStripMenuItem.ForeColor = Properties.Settings.Default.menuFontColor;
            ЗакрытьВсёToolStripMenuItem.ForeColor = Properties.Settings.Default.menuFontColor;
            ВыходToolStripMenuItem.ForeColor = Properties.Settings.Default.menuFontColor;
            ВыбратьТемуToolStripMenuItem.ForeColor = Properties.Settings.Default.menuFontColor;
            СветлаяToolStripMenuItem.ForeColor = Properties.Settings.Default.menuFontColor;
            ТёмнаяToolStripMenuItem.ForeColor = Properties.Settings.Default.menuFontColor;
            БирюзовыйФлексToolStripMenuItem.ForeColor = Properties.Settings.Default.menuFontColor;
            РозовыйКарнавалToolStripMenuItem.ForeColor = Properties.Settings.Default.menuFontColor;
            автосохранениеToolStripMenuItem.ForeColor = Properties.Settings.Default.menuFontColor;
            включитьToolStripMenuItem.ForeColor = Properties.Settings.Default.menuFontColor;
            отключитьToolStripMenuItem.ForeColor = Properties.Settings.Default.menuFontColor;
            секундToolStripMenuItem.ForeColor = Properties.Settings.Default.menuFontColor;
            минутToolStripMenuItem.ForeColor = Properties.Settings.Default.menuFontColor;
            минутToolStripMenuItem1.ForeColor = Properties.Settings.Default.menuFontColor;
            минутаToolStripMenuItem.ForeColor = Properties.Settings.Default.menuFontColor;
            минутыToolStripMenuItem.ForeColor = Properties.Settings.Default.menuFontColor;
            ЖирныйToolStripMenuItem1.ForeColor = Properties.Settings.Default.menuFontColor;
            СохранитьВсёtoolStripMenuItem1.ForeColor = Properties.Settings.Default.menuFontColor;
            КурсивToolStripMenuItem1.ForeColor = Properties.Settings.Default.menuFontColor;
            ПодчёркнутыйToolStripMenuItem1.ForeColor = Properties.Settings.Default.menuFontColor;
            ЗачёркнутыйToolStripMenuItem1.ForeColor = Properties.Settings.Default.menuFontColor;
            КурсивToolStripMenuItem.BackColor = Properties.Settings.Default.contextMenuColor;
            ЖирныйToolStripMenuItem.BackColor = Properties.Settings.Default.contextMenuColor;
            ПодчёркнутыйToolStripMenuItem.BackColor = Properties.Settings.Default.contextMenuColor;
            ЗачёркнутыйToolStripMenuItem.BackColor = Properties.Settings.Default.contextMenuColor;
        }

        /// <summary>
        /// Метод, меняющий стиль выбранного текста в активной вкладке.
        /// Тут в Visual Studio иногда не обрабатывается NullReferenceException, а в экзешнике обрабатывается.
        /// Если такое произошло, пожалуйста, проверь этот момент в экзешнике.
        /// </summary>
        /// <param name="fontStyle">Выбранный стиль текста.</param>

        private void ChangeFontStyle(FontStyle fontStyle)
        {
            Control control = tabControl.Controls[tabControl.SelectedIndex].Controls[0];
            RichTextBox richTextBox = control as RichTextBox;
            try
            {
                richTextBox.SelectionFont = new Font(richTextBox.Font, richTextBox.SelectionFont.Style ^ fontStyle);
            }
            catch (NullReferenceException)
            {
                MessageBox.Show("Нельзя менять стиль текста, где используются разные шрифты. " +
                    "Попробуйте изменить стиль текста с разными шрифтами раздельно.", "Внимание!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка! {ex.Message}", "Внимание!");
            }
        }

        /// <summary>
        /// Метод события, меняющий стиль выбранного текста на курсив.
        /// </summary>
        /// <param name="sender">Объект, над которым совершено событие.</param>
        /// <param name="e">Информация о событии.</param>

        private void КурсивToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ChangeFontStyle(FontStyle.Italic);
        }

        /// <summary>
        /// Метод события, меняющий стиль выбранного текста на жирный.
        /// </summary>
        /// <param name="sender">Объект, над которым совершено событие.</param>
        /// <param name="e">Информация о событии.</param>

        private void ЖирныйToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ChangeFontStyle(FontStyle.Bold);
        }

        /// <summary>
        /// Метод события, меняющий стиль выбранного текста на подчёркнутый.
        /// </summary>
        /// <param name="sender">Объект, над которым совершено событие.</param>
        /// <param name="e">Информация о событии.</param>

        private void ПодчёркнутыйToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ChangeFontStyle(FontStyle.Underline);
        }

        /// <summary>
        /// Метод события, меняющий стиль выбранного текста на зачёркнутый.
        /// </summary>
        /// <param name="sender">Объект, над которым совершено событие.</param>
        /// <param name="e">Информация о событии.</param>

        private void ЗачёркнутыйToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ChangeFontStyle(FontStyle.Strikeout);
        }

        /// <summary>
        /// Метод события, меняющий стиль выбранного текста на курсив.
        /// </summary>
        /// <param name="sender">Объект, над которым совершено событие.</param>
        /// <param name="e">Информация о событии.</param>

        private void КурсивToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (tabControl.TabPages.Count != 0)
            {
                try
                {
                    Control control = tabControl.Controls[tabControl.SelectedIndex].Controls[0];
                    RichTextBox richTextBox = control as RichTextBox;
                    richTextBox.SelectionFont = new Font(richTextBox.Font, richTextBox.SelectionFont.Style ^ FontStyle.Italic);
                }
                catch (NullReferenceException)
                {
                    MessageBox.Show("Нельзя менять стиль текста, где используются разные шрифты. " +
                        "Попробуйте изменить стиль текста с разными шрифтами раздельно.", "Внимание!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Произошла ошибка! {ex.Message}", "Внимание!");
                }
            }
        }

        /// <summary>
        /// Метод события, меняющий стиль выбранного текста на жирный.
        /// </summary>
        /// <param name="sender">Объект, над которым совершено событие.</param>
        /// <param name="e">Информация о событии.</param>

        private void ЖирныйToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (tabControl.TabPages.Count != 0)
            {
                try
                {
                    Control control = tabControl.Controls[tabControl.SelectedIndex].Controls[0];
                    RichTextBox richTextBox = control as RichTextBox;
                    richTextBox.SelectionFont = new Font(richTextBox.Font, richTextBox.SelectionFont.Style ^ FontStyle.Bold);
                }
                catch (NullReferenceException)
                {
                    MessageBox.Show("Нельзя менять стиль текста, где используются разные шрифты. " +
                        "Попробуйте изменить стиль текста с разными шрифтами раздельно.", "Внимание!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Произошла ошибка! {ex.Message}", "Внимание!");
                }
            }
        }

        /// <summary>
        /// Метод события, меняющий стиль выбранного текста на подчёркнутый.
        /// </summary>
        /// <param name="sender">Объект, над которым совершено событие.</param>
        /// <param name="e">Информация о событии.</param>

        private void ПодчёркнутыйToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (tabControl.TabPages.Count != 0)
            {
                try
                {
                    Control control = tabControl.Controls[tabControl.SelectedIndex].Controls[0];
                    RichTextBox richTextBox = control as RichTextBox;
                    richTextBox.SelectionFont = new Font(richTextBox.Font, richTextBox.SelectionFont.Style ^ FontStyle.Underline);
                }
                catch (NullReferenceException)
                {
                    MessageBox.Show("Нельзя менять стиль текста, где используются разные шрифты. " +
                        "Попробуйте изменить стиль текста с разными шрифтами раздельно.", "Внимание!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Произошла ошибка! {ex.Message}", "Внимание!");
                }
            }
        }

        /// <summary>
        /// Метод события, меняющий стиль выбранного текста на зачёркнутый.
        /// </summary>
        /// <param name="sender">Объект, над которым совершено событие.</param>
        /// <param name="e">Информация о событии.</param>

        private void ЗачёркнутыйToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (tabControl.TabPages.Count != 0)
            {
                try
                {
                    Control control = tabControl.Controls[tabControl.SelectedIndex].Controls[0];
                    RichTextBox richTextBox = control as RichTextBox;
                    richTextBox.SelectionFont = new Font(richTextBox.Font, richTextBox.SelectionFont.Style ^ FontStyle.Strikeout);
                }
                catch (NullReferenceException)
                {
                    MessageBox.Show("Нельзя менять стиль текста, где используются разные шрифты. " +
                        "Попробуйте изменить стиль текста с разными шрифтами раздельно.", "Внимание!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Произошла ошибка! {ex.Message}", "Внимание!");
                }
            }
        }

        /// <summary>
        /// Метод события, меняющий шрифт всего текста активной вкладки на выбранный пользователем.
        /// </summary>
        /// <param name="sender">Объект, над которым совершено событие.</param>
        /// <param name="e">Информация о событии.</param>

        private void ФорматToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (tabControl.TabPages.Count != 0)
                {
                    Control control = tabControl.Controls[tabControl.SelectedIndex].Controls[0];
                    RichTextBox richTextBox = control as RichTextBox;
                    FontDialog fontDialog = new FontDialog();
                    if (fontDialog.ShowDialog() != DialogResult.Cancel)
                    {
                        richTextBox.Font = fontDialog.Font;
                        richTextBox.ForeColor = fontDialog.Color;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка! {ex.Message}", "Внимание!");
            }
        }

        /// <summary>
        /// Метод события, меняющий интервал таймера на 30 секунд и начинающий его работу.
        /// </summary>
        /// <param name="sender">Объект, над которым совершено событие.</param>
        /// <param name="e">Информация о событии.</param>

        private void ТридцатьСекундToolStripMenuItem_Click(object sender, EventArgs e)
        {
            autoSaveTimer.Interval = 30000;
            autoSaveTimer.Enabled = true;
            Properties.Settings.Default.autoSaveNum = autoSaveTimer.Interval;
            Properties.Settings.Default.autoSaveOn = autoSaveTimer.Enabled;
            Properties.Settings.Default.Save();
        }

        /// <summary>
        /// Метод события, меняющий интервал таймера на 1 минуту и начинающий его работу.
        /// </summary>
        /// <param name="sender">Объект, над которым совершено событие.</param>
        /// <param name="e">Информация о событии.</param>

        private void ОднаМинутаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            autoSaveTimer.Interval = 60000;
            autoSaveTimer.Enabled = true;
            Properties.Settings.Default.autoSaveNum = autoSaveTimer.Interval;
            Properties.Settings.Default.autoSaveOn = autoSaveTimer.Enabled;
            Properties.Settings.Default.Save();
        }

        /// <summary>
        /// Метод события, меняющий интервал таймера на 3 минуты и начинающий его работу.
        /// </summary>
        /// <param name="sender">Объект, над которым совершено событие.</param>
        /// <param name="e">Информация о событии.</param>

        private void ТриМинутыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            autoSaveTimer.Interval = 180000;
            autoSaveTimer.Enabled = true;
            Properties.Settings.Default.autoSaveNum = autoSaveTimer.Interval;
            Properties.Settings.Default.autoSaveOn = autoSaveTimer.Enabled;
            Properties.Settings.Default.Save();
        }

        /// <summary>
        /// Метод события, меняющий интервал таймера на 5 минут и начинающий его работу.
        /// </summary>
        /// <param name="sender">Объект, над которым совершено событие.</param>
        /// <param name="e">Информация о событии.</param>

        private void ПятьМинутToolStripMenuItem_Click(object sender, EventArgs e)
        {
            autoSaveTimer.Interval = 300000;
            autoSaveTimer.Enabled = true;
            Properties.Settings.Default.autoSaveNum = autoSaveTimer.Interval;
            Properties.Settings.Default.autoSaveOn = autoSaveTimer.Enabled;
            Properties.Settings.Default.Save();
        }

        /// <summary>
        /// Метод события, меняющий интервал таймера на 10 минут и начинающий его работу.
        /// </summary>
        /// <param name="sender">Объект, над которым совершено событие.</param>
        /// <param name="e">Информация о событии.</param>

        private void ДесятьМинутToolStripMenuItem_Click(object sender, EventArgs e)
        {
            autoSaveTimer.Interval = 600000;
            autoSaveTimer.Enabled = true;
            Properties.Settings.Default.autoSaveNum = autoSaveTimer.Interval;
            Properties.Settings.Default.autoSaveOn = autoSaveTimer.Enabled;
            Properties.Settings.Default.Save();
        }

        /// <summary>
        /// Метод, останавливающий работу таймера.
        /// </summary>
        /// <param name="sender">Объект, над которым совершено событие.</param>
        /// <param name="e">Информация о событии.</param>

        private void ОтключитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            autoSaveTimer.Enabled = false;
            Properties.Settings.Default.autoSaveOn = autoSaveTimer.Enabled;
            Properties.Settings.Default.Save();
        }

        /// <summary>
        /// Метод события, сохраняющий все файлы во вкладках или выводящий сообщение об ошибке, если сохранить не получилось.
        /// </summary>
        /// <param name="sender">Объект, над которым совершено событие.</param>
        /// <param name="e">Информация о событии.</param>

        private void AutoSaveTimer_Tick(object sender, EventArgs e)
        {
            for (int i = 0; i < tabControl.TabPages.Count; i++)
            {
                try
                {
                    string realName = tabControl.TabPages[i].Name;
                    if (tabControl.TabPages[i].Name[tabControl.TabPages[i].Name.Length - 1] == '*')
                        realName = tabControl.TabPages[i].Name.Substring(0, tabControl.TabPages[i].Name.Length - 1);
                    if (File.Exists(realName))
                    {
                        Control control = tabControl.Controls[i].Controls[0];
                        RichTextBox richTextBox = control as RichTextBox;
                        if (Path.GetExtension(realName) == ".rtf")
                            richTextBox.SaveFile(realName);
                        else
                        {
                            StreamWriter sw = new StreamWriter(realName);
                            sw.Write(richTextBox.Text);
                            sw.Close();
                        }
                        tabControl.TabPages[i].Name = realName;
                    }
                    else
                    {
                        MessageBox.Show($"Файла {realName} больше не существует, не удалось сохранить.", "Внимание!");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Произошла ошибка при автосохранении! {ex.Message}", "Внимание!");
                }
            }
        }

        /// <summary>
        /// Метод события закрытия формы, спрашивающий, нужно ли сохранить измененные файлы.
        /// </summary>
        /// <param name="sender">Объект, над которым совершено событие.</param>
        /// <param name="e">Информация о событии.</param>

        private void Notepad_FormClosing(object sender, FormClosingEventArgs e)
        {
            for (int i = 0; i < tabControl.TabPages.Count; i++)
            {
                try
                {
                    if (tabControl.TabPages[i].Name[tabControl.TabPages[i].Name.Length - 1] == '*')
                    {
                        DialogResult result = MessageBox.Show($"Сохранить изменения в " +
                            $"{tabControl.TabPages[i].Name.Substring(0, tabControl.TabPages[i].Name.Length - 1)}?", "Внимание!",
                            MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            if (File.Exists(tabControl.TabPages[i].Name.Substring(0, tabControl.TabPages[i].Name.Length - 1)))
                            {
                                Control control = tabControl.Controls[i].Controls[0];
                                RichTextBox richTextBox = control as RichTextBox;
                                if (Path.GetExtension(tabControl.TabPages[i].Name.Substring(0, tabControl.TabPages[i].Name.Length - 1)) == ".rtf")
                                    richTextBox.SaveFile(tabControl.TabPages[i].Name.Substring(0, tabControl.TabPages[i].Name.Length - 1));
                                else
                                {
                                    StreamWriter sw = new StreamWriter(tabControl.TabPages[i].Name.Substring(0, tabControl.TabPages[i].Name.Length - 1));
                                    sw.Write(richTextBox.Text);
                                    sw.Close();
                                }
                                tabControl.TabPages[i].Name = tabControl.TabPages[i].Name.Substring(0, tabControl.TabPages[i].Name.Length - 1);
                            }
                            else
                                MessageBox.Show($"Файла {tabControl.TabPages[i].Name.Substring(0, tabControl.TabPages[i].Name.Length - 1)} больше не существует, попробуйте воспользоваться функцией \"Сохранить как\".", "Внимание!");
                        }
                        else if (result == DialogResult.No)
                            return;
                        else e.Cancel = true;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Произошла ошибка! {ex.Message}", "Внимание!");
                }
            }
        }

        /// <summary>
        /// Метод, позволяющий понять, изменялся ли текст во вкладке, добавляя к ее имени звездочку.
        /// </summary>
        /// <param name="sender">Объект, над которым совершено событие.</param>
        /// <param name="e">Информация о событии.</param>

        private void TextChangedHandler(object sender, EventArgs e)
        {
            try
            {
                RichTextBox richTextBox = (RichTextBox)sender;
                TabPage tabPage = (TabPage)richTextBox.Parent;
                if (tabPage.Name[tabPage.Name.Length - 1] != '*')
                    tabPage.Name += "*";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка! {ex.Message}", "Внимание!");
            }
        }
    }

    /// <summary>
    /// Класс рендера, позволяющий менять цвет вкладок вторичного меню на голубой. Если вкладка меню была выбрана, 
    /// её цвет также меняется на другой.
    /// </summary>

    public class GradientInactiveCaptionRenderer : ToolStripProfessionalRenderer
    {
        protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e)
        {
            Rectangle rectangle = new Rectangle(Point.Empty, e.Item.Size);
            Color color = e.Item.Selected ? Color.AliceBlue : SystemColors.GradientInactiveCaption;
            using (SolidBrush brush = new SolidBrush(color))
                e.Graphics.FillRectangle(brush, rectangle);
        }
    }

    /// <summary>
    /// Класс рендера, позволяющий менять цвет вкладок вторичного меню на серый. Если вкладка меню была выбрана,
    /// её цвет также меняется на другой.
    /// </summary>

    public class WindowFrameRenderer : ToolStripProfessionalRenderer
    {
        protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e)
        {
            Rectangle rectangle = new Rectangle(Point.Empty, e.Item.Size);
            Color color = e.Item.Selected ? SystemColors.ControlDark : SystemColors.WindowFrame;
            using (SolidBrush brush = new SolidBrush(color))
                e.Graphics.FillRectangle(brush, rectangle);
        }
    }

    /// <summary>
    /// Класс рендера, позволяющий менять цвет вкладок вторичного меню на бирюзовый. Если вкладка меню была выбрана, 
    /// её цвет также меняется на другой.
    /// </summary>

    public class DarkCyanRenderer : ToolStripProfessionalRenderer
    {
        protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e)
        {
            Rectangle rectangle = new Rectangle(Point.Empty, e.Item.Size);
            Color color = e.Item.Selected ? Color.CadetBlue : Color.DarkCyan;
            using (SolidBrush brush = new SolidBrush(color))
                e.Graphics.FillRectangle(brush, rectangle);
        }
    }

    /// <summary>
    /// Класс рендера, позволяющий менять цвет вкладок вторичного меню на розовый. Если вкладка меню была выбрана, 
    /// её цвет также меняется на другой.
    /// </summary>

    public class MediumVioletRedRenderer : ToolStripProfessionalRenderer
    {
        protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e)
        {
            Rectangle rectangle = new Rectangle(Point.Empty, e.Item.Size);
            Color color = e.Item.Selected ? Color.HotPink : Color.MediumVioletRed;
            using (SolidBrush brush = new SolidBrush(color))
                e.Graphics.FillRectangle(brush, rectangle);
        }
    }
}
