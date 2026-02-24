using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Windows.Forms;
using Microsoft.Win32;

namespace FortyOne.AudioSwitcher.Helpers
{
    public class VolumeOSD : Form
    {
        private static VolumeOSD _instance;
        private readonly Timer _hideTimer;
        private readonly Timer _fadeTimer;

        private int _volume;
        private bool _isMuted;
        private float _targetOpacity;
        private const float MAX_OPACITY = 0.95f;

        // ── Геометрия ───────────────────────────
        private const int W = 100;
        private const int H = 200;
        private const int BAR_WIDTH = 20;
        private const int THUMB_HEIGHT = 20;

        private VolumeOSD()
        {
            // Настройка окна
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.Manual;
            ShowInTaskbar = false;
            TopMost = true;
            DoubleBuffered = true;
            BackColor = Color.FromArgb(0, 0, 0); // Тёмный фон формы
            
            // Фиксируем размер, чтобы избежать мерцания и изменения размера
            Size = new Size(W, H);
            MinimumSize = new Size(W, H);
            MaximumSize = new Size(W, H);
            
            Opacity = 0;

            // Позиция: левый верхний угол с заданным отступом
            Location = new Point(80, 80);

            // Таймер скрытия
            _hideTimer = new Timer { Interval = 2000 };
            _hideTimer.Tick += (sender, args) => { _targetOpacity = 0; _fadeTimer.Start(); _hideTimer.Stop(); };

            // Таймер анимации появления/исчезновения
            _fadeTimer = new Timer { Interval = 16 };
            _fadeTimer.Tick += (sender, args) =>
            {
                if (Math.Abs(_targetOpacity - Opacity) < 0.05)
                {
                    Opacity = _targetOpacity;
                    _fadeTimer.Stop();
                    if (Opacity == 0) Hide();
                }
                else Opacity += (_targetOpacity - (float)Opacity) * 0.25f;
            };
        }

        public static void ShowVolume(int percent, bool isMuted)
        {
            if (_instance == null || _instance.IsDisposed) _instance = new VolumeOSD();

            _instance._volume = percent < 0 ? 0 : (percent > 100 ? 100 : percent);
            _instance._isMuted = isMuted;
            _instance._targetOpacity = MAX_OPACITY;

            _instance.Show();
            _instance.Invalidate();
            _instance._fadeTimer.Start();
            _instance._hideTimer.Stop();
            _instance._hideTimer.Start();
        }

        private Color GetAccentColor()
        {
            try
            {
                // Берем системный акцентный цвет Windows
                object colorValueObj = Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\DWM", "ColorizationColor", 0);
                if (colorValueObj != null)
                {
                    int colorValue = (int)colorValueObj;
                    // Извлекаем RGB
                    return Color.FromArgb(255, (colorValue >> 16) & 0xFF, (colorValue >> 8) & 0xFF, colorValue & 0xFF);
                }
            }
            catch { }
            return Color.Teal; // Запасной вариант
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.None; // Пиксельная четкость
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

            int cx = Width / 2;
            Color accent = GetAccentColor();

            // ── Отрисовка числа / Крестика ───────────
            using (var font = new Font("Segoe UI", 12, FontStyle.Regular))
            {
                string txt = _volume.ToString();
                var sz = g.MeasureString(txt, font);
                int tx = cx - (int)(sz.Width / 2) + 1;
                int ty = Height - 38;

                if (_isMuted || _volume <= 0)
                {
                    int crossSize = 18;
                    int kx = cx - crossSize / 2;
                    int ky = ty + (int)sz.Height / 2 - crossSize / 2 - 2;

                    using (var p = new Pen(Color.White, 3))
                    {
                        g.SmoothingMode = SmoothingMode.AntiAlias;
                        g.DrawLine(p, kx, ky, kx + crossSize, ky + crossSize);
                        g.DrawLine(p, kx + crossSize, ky, kx, ky + crossSize);
                        g.SmoothingMode = SmoothingMode.None;
                    }
                }
                else
                {
                    g.DrawString(txt, font, Brushes.White, tx, ty);
                }
            }

            // ── Параметры полоски ──────────────────
            int barX = cx - BAR_WIDTH / 2;
            int barTop = 15;
            int barBottom = Height - 48;
            int barFullH = barBottom - barTop;

            // Расчитываем диапазон движения для КВАДРАТНОГО ползунка
            // Ползунок должен перемещаться от barTop до (barBottom - THUMB_HEIGHT)
            int thumbRange = barFullH - THUMB_HEIGHT;
            int thumbY = barBottom - THUMB_HEIGHT - (int)(thumbRange * (_volume / 100.0));
            
            // Высота заполнения (доходит до низа ползунка)
            int fillH = barBottom - (thumbY + THUMB_HEIGHT);

            // 1. Фон полоски (темно-серый)
            using (var b = new SolidBrush(Color.FromArgb(75, 75, 75)))
                g.FillRectangle(b, barX, barTop, BAR_WIDTH, barFullH);

            // 2. Заполнение (Акцентный цвет)
            if (fillH > 0)
            {
                using (var b = new SolidBrush(accent))
                    g.FillRectangle(b, barX, thumbY + THUMB_HEIGHT, BAR_WIDTH, fillH);
            }

            // 3. Белый ползунок (Thumb)
            g.FillRectangle(Brushes.White, barX, thumbY, BAR_WIDTH, THUMB_HEIGHT);
        }

        // Делаем окно "сквозным" для кликов и прозрачным для системы
        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.ExStyle |= 0x80000; // WS_EX_LAYERED
                cp.ExStyle |= 0x20;    // WS_EX_TRANSPARENT (click-through)
                return cp;
            }
        }

        protected override bool ShowWithoutActivation
        {
            get { return true; }
        }
    }
}
