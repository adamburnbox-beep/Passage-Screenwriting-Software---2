namespace Passage.Web.Services;

/// <summary>
/// Deals the input for an Ideation burst round, offline, so a sitting can
/// start from nothing. The worksheet's lanes mostly say "pick blind" or
/// "don't cherry-pick"; a writer choosing their own seed is the deliberating
/// the drill forbids, and a writer with nothing has no seed at all. The
/// banks are small and plain on purpose: the seed is meant to be forced on
/// you, not admired. Nothing here touches a script or the network.
/// </summary>
public static class IdeationDealer
{
    /// <summary>One dealt input for the lane (1–10). Never empty.</summary>
    public static string Deal(int lane, Random? random = null)
    {
        var rng = random ?? Random.Shared;
        return lane switch
        {
            1 => NoiseSeed(rng),
            2 => Pick(Grievances, rng),
            3 => Pick(Catastrophes, rng),
            4 => Collision(rng),
            5 => Pick(Nouns, rng),
            6 => $"{Pick(People, rng)} — who won't say {Pick(Withheld, rng)}.",
            7 => Pick(People, rng),
            8 => Pick(StockTypes, rng),
            9 => TwoVoices(rng),
            _ => Pick(Flaws, rng),
        };
    }

    private static string NoiseSeed(Random rng) => rng.Next(3) switch
    {
        0 => "Object: " + Pick(Objects, rng),
        1 => "Headline: " + Pick(Headlines, rng),
        _ => "Image: " + Pick(Images, rng),
    };

    private static string Collision(Random rng)
    {
        var (one, other) = PickTwo(Fragments, rng);
        return $"{one}  ×  {other}";
    }

    private static string TwoVoices(Random rng)
    {
        var (first, second) = PickTwo(People, rng);
        return $"{first} and {second}, {Pick(Places, rng)}.";
    }

    private static string Pick(string[] bank, Random rng) => bank[rng.Next(bank.Length)];

    // Two different entries from one bank, for the lanes that collide things.
    private static (string, string) PickTwo(string[] bank, Random rng)
    {
        var first = rng.Next(bank.Length);
        var second = (first + 1 + rng.Next(bank.Length - 1)) % bank.Length;
        return (bank[first], bank[second]);
    }

    private static readonly string[] Objects =
    {
        "a hotel key card with the magnetic strip scratched off",
        "a child's shoe, one, on a bus seat",
        "a wedding ring in a jar of screws",
        "a library book forty years overdue",
        "a folding chair chained to a lamppost",
        "a violin case with no violin in it",
        "a receipt for one coffin",
        "a set of house keys in the freezer",
        "a passport with the photo cut out",
        "a birthday cake nobody has cut",
        "a dog lead with no dog on the end",
        "a hearing aid on a restaurant table",
        "a bag of soil on the back seat of a taxi",
        "a school trophy in a charity shop window",
        "a lighthouse bulb, boxed, on a doorstep",
        "a phone that only rings at 3 a.m.",
        "a chessboard with both kings missing",
        "a wheelbarrow full of letters",
        "a lifejacket hanging in a city flat",
        "a bell with the clapper removed",
    };

    private static readonly string[] Headlines =
    {
        "COUNCIL VOTES TO RENAME THE RIVER",
        "LAST FERRY CANCELLED, NO REPLACEMENT PLANNED",
        "LOCAL MAN RETURNS LIBRARY BOOK AFTER 61 YEARS",
        "SCHOOL TO CLOSE FOR ONE DAY, REASON WITHHELD",
        "MISSING LIGHTHOUSE KEEPER FOUND IN OWN KITCHEN",
        "TOWN RUNS OUT OF SALT",
        "ENTIRE STREET SELLS UP ON THE SAME DAY",
        "FUNERAL HOME OPENS CAFÉ",
        "CHOIR BANNED FROM SINGING INDOORS",
        "BRIDGE TO BE DISMANTLED AND SOLD BY THE BOLT",
        "WOMAN INHERITS ZOO SHE DIDN'T KNOW EXISTED",
        "MAYOR'S TWIN DENIES BEING MAYOR",
        "LOTTERY WINNER STILL CLOCKING IN AT THE FACTORY",
        "FLOOD UNCOVERS SECOND CEMETERY",
        "VILLAGE VOTES TO STOP THE CLOCK",
        "BAKER WINS CASE AGAINST HIS OWN SON",
        "ALL THE TOWN'S DOGS RAN WEST ON TUESDAY",
        "SCHOOL BUS DRIVER RETIRES AFTER 44 YEARS, TAKES BUS",
    };

    private static readonly string[] Images =
    {
        "a wedding dress on a washing line in the rain",
        "a man in a suit asleep on a roundabout",
        "two hearses parked nose to nose",
        "a child directing traffic, and the traffic obeying",
        "a dining table set for twelve in a field",
        "a woman reading a letter to a horse",
        "a piano halfway up a staircase, abandoned",
        "a queue outside a shop that has clearly closed for good",
        "a lit Christmas tree in a window in July",
        "a bride running for a bus",
        "a funeral where everyone is laughing",
        "a boat on a trailer, a hundred miles from any water",
        "a crowd on a beach all facing away from the sea",
        "a doctor smoking outside a hospital, crying",
        "a house with every window bricked up but one",
        "a boy carrying a door across a car park",
        "a dog waiting outside a courthouse",
        "an empty swimming pool with a chair at the bottom",
    };

    private static readonly string[] Grievances =
    {
        "Someone else took the credit, and thanking them was your job.",
        "A rule applied to you was waived for them, in front of you.",
        "An apology that was really an explanation of why you were wrong.",
        "You were right, early, and it cost you more than being wrong would have.",
        "They asked for honesty and punished it.",
        "Your name was left off, and nobody noticed but you.",
        "The thing you gave up for them, they never wanted.",
        "A promise kept to the letter and broken in every other way.",
        "Being talked about as if you weren't in the room.",
        "A kindness done to you so it could be mentioned later.",
        "They forgot, and you'll never be able to prove it mattered.",
        "The fix was easy and they made you beg for it anyway.",
        "You were consulted after the decision.",
        "Being the only one who remembers how it actually happened.",
        "A door held open for everyone else and let go for you.",
        "Watching them get away with it, and being told to be gracious.",
        "The queue you waited in was for nothing.",
        "Being thanked for something you were forced to do.",
    };

    private static readonly string[] Catastrophes =
    {
        "The only bridge into town closes for good at midnight.",
        "The will is read and everything goes to the dog.",
        "The groom's twin turns up at the wedding, uninvited, in the same suit.",
        "Every phone in the building rings at once, then stops.",
        "The school burns down the night before the exam.",
        "A stranger returns your late mother's coat.",
        "The tide goes out and doesn't come back.",
        "Your understudy is better than you and everyone has seen it.",
        "The factory's last shift is announced by text.",
        "The confession is found in the wrong envelope.",
        "The lift stops between floors with the two of them inside.",
        "The lottery ticket is in the coat you gave away.",
        "The village well runs dry on the day of the fête.",
        "The witness recognises the judge.",
        "The baby is born on the ferry, and the ferry is turning back.",
        "Every clock in the house is found stopped at the same time.",
        "The boss's funeral and the boss's wedding fall on the same day.",
        "The dog comes home. The child doesn't.",
    };

    private static readonly string[] Fragments =
    {
        "a marriage proposal",
        "a bomb-disposal briefing",
        "a school nativity",
        "a bailiff at the door",
        "a lifeboat launch",
        "a job interview",
        "a driving test",
        "a hostage negotiation",
        "the reading of a will",
        "a cake competition",
        "a fire drill",
        "a confession box",
        "the last bus of the night",
        "a first date",
        "a hospice bedside",
        "a chess tournament",
        "a stalled lift",
        "a sheep auction",
        "a parole hearing",
        "a child's swimming lesson",
        "a séance",
        "an eviction",
        "a citizenship ceremony",
        "a card game for money",
        "a lighthouse in fog",
        "a karaoke night",
        "a hunger strike",
        "a border crossing",
        "a driving lesson",
        "an organ transplant call",
    };

    private static readonly string[] Nouns =
    {
        "a lighthouse", "a tourniquet", "a hinge", "a fuse", "a compass",
        "a scaffold", "a dam", "an anchor", "a mirror", "a splinter",
        "a bridge", "a kettle", "a scar", "a lock", "a net",
        "a ladder", "a wick", "a fence", "a hive", "a ledger",
        "a tide", "a magnet", "a bandage", "a well", "a kite",
        "a knot", "a siren", "a seed", "a fault line", "a glove",
    };

    private static readonly string[] People =
    {
        "a retired midwife",
        "a night-shift security guard",
        "a driving instructor on her last day",
        "a priest who has lost his voice",
        "a lottery winner still doing his old job",
        "a child who has just learned to lie",
        "a bailiff who used to be a nurse",
        "a lighthouse keeper's widow",
        "a football referee, off duty",
        "a locksmith who has never been robbed",
        "an understudy who has never gone on",
        "a translator at a funeral",
        "a chef who can no longer taste",
        "a magistrate on her first morning",
        "a bus driver who knows every passenger's name",
        "a taxidermist in mourning",
        "a wedding photographer at his own divorce",
        "a stepfather meeting the stepchild's real father",
        "a pilot who is afraid of lifts",
        "a nun who has just won an argument",
        "a debt collector who is owed money",
        "a school caretaker who was once a pupil there",
        "a woman who has been declared dead by mistake",
        "a twin who has never been mistaken for the other",
    };

    private static readonly string[] Withheld =
    {
        "why they stopped",
        "where the money went",
        "who was in the car",
        "what was in the letter",
        "why they came back",
        "who they were with that night",
        "what they promised the dying man",
        "why they never learned to drive",
        "what they said to make the child cry",
        "why the door is always locked",
        "what happened to the first one",
        "why they changed their name",
        "who taught them to do that",
        "what the doctor actually said",
        "why they won't go to the coast",
        "what they were doing in the church",
    };

    private static readonly string[] StockTypes =
    {
        "the villain", "the sidekick", "the ingenue", "the mentor",
        "the henchman", "the femme fatale", "the comic relief", "the wise fool",
        "the reluctant hero", "the busybody neighbour", "the corrupt official",
        "the faithful servant", "the mad scientist", "the prodigal child",
        "the stern matriarch", "the drunk uncle", "the innocent bystander",
        "the loyal dog, given speech", "the jaded detective", "the young pretender",
        "the grieving widow", "the braggart soldier", "the crooked lawyer",
        "the girl next door", "the tyrant", "the hermit", "the trickster",
        "the damsel who is not in distress",
    };

    private static readonly string[] Places =
    {
        "in a stalled lift",
        "on the last bus",
        "at a bus stop in the rain",
        "in a hospital corridor at 4 a.m.",
        "in the queue at the post office",
        "on a roof, at night",
        "in a car that won't start",
        "at a wake, in the kitchen",
        "on a ferry that has turned back",
        "in a waiting room with one chair",
        "in the changing rooms after the match",
        "at the back of a wedding",
        "in a lighthouse, in fog",
        "on hold, on speakerphone",
        "at a border crossing",
        "on the stairs, one going up, one coming down",
    };

    private static readonly string[] Flaws =
    {
        "Can't let anyone else finish a sentence.",
        "Says yes to everything and means none of it.",
        "Would rather be right than be there.",
        "Mistakes being needed for being loved.",
        "Apologises before anyone has accused them.",
        "Keeps score, and shows the scorecard.",
        "Cannot ask for help, only wait for it.",
        "Tells the truth as a weapon.",
        "Lies to make the room comfortable.",
        "Treats every kindness as a debt to be repaid immediately.",
        "Punishes the messenger, every time.",
        "Has to be the one who leaves first.",
        "Mistakes silence for agreement.",
        "Cannot bear to be ordinary at anything.",
        "Confuses forgiving with forgetting, and does neither.",
        "Only trusts people who have been hurt as badly.",
        "Rescues people who have not asked to be rescued.",
        "Believes loyalty should never be tested, then tests it.",
        "Hoards small secrets against the day they might be useful.",
        "Would rather lose alone than win with help.",
    };
}
