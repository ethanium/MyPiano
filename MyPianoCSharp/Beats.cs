namespace MyPiano;

/// <summary>
/// A rhythm pattern (port of MidiBeat.bas). Sections are: 0 intro, 1 main loop, 2 fill 1, 3 fill 2, 4 ending.
/// Each section is a list of steps; each step is a list of percussion note numbers played together.
/// </summary>
public sealed class Beat
{
    public string Name { get; }
    public int Rate { get; }
    public int[][][] Sections { get; }

    public Beat(string name, int rate, string data)
    {
        Name = name;
        Rate = rate;
        Sections = data.Split(';')
            .Select(section => section.Split('|')
                .Select(step => step.Split(',', StringSplitOptions.RemoveEmptyEntries)
                    .Select(int.Parse).ToArray())
                .ToArray())
            .ToArray();
    }

    public static readonly Beat[] All =
    {
        new("Disco", 170,
            "44,35,55|46|44,40,35|42;" +
            "44,35|46|44,40,35|42;" +
            "44,35|44,35,40|44,35,40|44,35,40|44,35,55|46|44,40,35|42;" +
            "44,35|44,40|44,40|44,35;" +
            "44,35,55"),

        new("Rock", 150,
            "44,35,55||44||44,40||44|;" +
            "44,35||44,35||44,40||44||44,35||44||44,40||44|;" +
            "44,50,35|50|44,48,35|48|44,47,35|47|44,45,35|45|44,35,55||44||44,40||44|;" +
            "44,35||46||40,44||42|;" +
            "44||44||44||44,35,55"),

        new("16 Beat", 100,
            "44|44|44|44|44|44|44|44;" +
            "44,35|44|44|44|44,40|44|44|44;" +
            "44,40|44,40|44,40|44,40|44,40|44,40|44,40|44,40|44,35,55|44|44|44|44,40|44|44|44;" +
            "44,35|44|44|46|40|42|44|44;" +
            "44,35,55"),
    };
}
