namespace TruckExcelAssistant;

internal static class AppTheme
{
    public static readonly Color WindowBackground = Color.FromArgb(244, 247, 250);
    public static readonly Color Surface = Color.White;
    public static readonly Color Sidebar = Color.FromArgb(24, 60, 54);
    public static readonly Color SidebarHover = Color.FromArgb(38, 82, 73);
    public static readonly Color Accent = Color.FromArgb(35, 106, 91);
    public static readonly Color AccentSoft = Color.FromArgb(228, 242, 238);
    public static readonly Color TextPrimary = Color.FromArgb(23, 32, 51);
    public static readonly Color TextSecondary = Color.FromArgb(92, 104, 120);
    public static readonly Color Border = Color.FromArgb(216, 222, 231);
    public static readonly Color InputBorder = Color.FromArgb(190, 199, 210);
    public static readonly Color Warning = Color.FromArgb(166, 91, 0);
    public static readonly Color Miguno = Color.FromArgb(176, 126, 0);
    public static readonly Color Agrico = Color.FromArgb(50, 116, 55);

    public static Button CreatePrimaryButton(string text)
    {
        return new Button
        {
            Text = text,
            AutoSize = false,
            Height = 40,
            FlatStyle = FlatStyle.Flat,
            BackColor = Accent,
            ForeColor = Color.White,
            Cursor = Cursors.Hand,
            Font = new Font("Segoe UI", 9F, FontStyle.Bold),
            Margin = new Padding(0, 4, 0, 0),
            UseVisualStyleBackColor = false
        }.WithFlatBorder(Accent);
    }

    public static Button CreateSecondaryButton(string text)
    {
        return new Button
        {
            Text = text,
            AutoSize = false,
            Height = 38,
            FlatStyle = FlatStyle.Flat,
            BackColor = Surface,
            ForeColor = TextPrimary,
            Cursor = Cursors.Hand,
            Margin = new Padding(0, 4, 0, 0),
            UseVisualStyleBackColor = false
        }.WithFlatBorder(InputBorder);
    }

    private static Button WithFlatBorder(this Button button, Color color)
    {
        button.FlatAppearance.BorderColor = color;
        button.FlatAppearance.BorderSize = 1;
        button.FlatAppearance.MouseOverBackColor = button.BackColor == Accent
            ? Color.FromArgb(29, 91, 78)
            : Color.FromArgb(245, 247, 250);
        return button;
    }
}
