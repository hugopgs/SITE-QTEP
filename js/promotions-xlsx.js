// Lecture du fichier Suivi_promotion.xlsx pour la page "Suivi des promotions".
// Pour mettre à jour le site : modifier le fichier Excel et le remettre en ligne, rien d'autre.
//
// Format attendu du fichier Excel :
// - un onglet par promotion, nommé "Promotion 2025", "Promotion 2026", ...
// - colonnes dans cet ordre (la ligne 1 contient les titres) :
//   A NOM | B PRENOM | C STAGE M2 (intitulé) | D INSTITUTION | E VILLE - PAYS |
//   F POSTE ACTUEL | G INSTITUTION | H VILLE - PAYS | I DIRECTEUR THESE
// - "VILLE - PAYS" s'écrit "Ville - Pays" (ex : "Esch-sur-Alzette - Luxembourg" ou "Grenoble, France")

const XLSX_FILE = '../data/Suivi_promotion.xlsx';

// Noms des promotions (facultatif : sans nom, on affiche "Promotion 2027")
const PROMOTION_NAMES = {
    "2025": "Promotion Mikhail Lukin"
};

// Couleurs attribuées aux promotions, de la plus récente à la plus ancienne
const PROMOTION_PALETTE = ['#3498db', '#e74c3c', '#2ecc71', '#9b59b6', '#f39c12', '#16a085', '#d35400', '#34495e'];

const JOB_TYPES = {
    phd: "Doctorant·e",
    postdoc: "Post-doc",
    engineer: "Ingénieur·e",
    researcher: "Chercheur·se",
    consultant: "Consultant·e",
    other: "Autre"
};

// Reconnaissance du type de poste à partir du texte de la colonne "POSTE ACTUEL"
const JOB_KEYWORDS = [
    ['phd', /doctorant|doctorante|ph\.? ?d|thèse|these/i],
    ['postdoc', /post-?doc/i],
    ['engineer', /ing[ée]nieur|engineer/i],
    ['researcher', /chercheur|chercheuse|researcher|scientist/i],
    ['consultant', /consult/i]
];

// Noms de pays harmonisés en français
const COUNTRY_NAMES = {
    austria: 'Autriche', autriche: 'Autriche',
    switzerland: 'Suisse', suisse: 'Suisse',
    denmark: 'Danemark', danemark: 'Danemark',
    netherlands: 'Pays-Bas', 'the netherlands': 'Pays-Bas', 'pays-bas': 'Pays-Bas', 'pays bas': 'Pays-Bas',
    germany: 'Allemagne', allemagne: 'Allemagne',
    'united kingdom': 'Royaume-Uni', uk: 'Royaume-Uni', 'royaume-uni': 'Royaume-Uni', england: 'Royaume-Uni',
    spain: 'Espagne', espagne: 'Espagne',
    italy: 'Italie', italie: 'Italie',
    belgium: 'Belgique', belgique: 'Belgique',
    sweden: 'Suède', suede: 'Suède',
    finland: 'Finlande', finlande: 'Finlande',
    norway: 'Norvège', norvege: 'Norvège',
    luxembourg: 'Luxembourg',
    france: 'France',
    usa: 'États-Unis', 'united states': 'États-Unis', 'etats-unis': 'États-Unis',
    canada: 'Canada', japan: 'Japon', japon: 'Japon', singapore: 'Singapour', singapour: 'Singapour',
    australia: 'Australie', australie: 'Australie', israel: 'Israël', poland: 'Pologne', pologne: 'Pologne',
    'czech republic': 'Tchéquie', czechia: 'Tchéquie', portugal: 'Portugal', ireland: 'Irlande', irlande: 'Irlande'
};

// Coordonnées des villes déjà connues (les autres sont cherchées automatiquement sur OpenStreetMap)
const KNOWN_CITIES = {
    paris: [48.8466, 2.3447], palaiseau: [48.7146, 2.2110], saclay: [48.7120, 2.1700], orsay: [48.6980, 2.1870],
    'gif-sur-yvette': [48.7090, 2.1360], grenoble: [45.1885, 5.7245], strasbourg: [48.5734, 7.7521],
    lyon: [45.7640, 4.8357], marseille: [43.2965, 5.3698], toulouse: [43.6047, 1.4442], bordeaux: [44.8378, -0.5792],
    lille: [50.6292, 3.0573], nice: [43.7102, 7.2620], rennes: [48.1173, -1.6778], nantes: [47.2184, -1.5536],
    montpellier: [43.6108, 3.8767], besancon: [47.2378, 6.0241], dijon: [47.3220, 5.0415], caen: [49.1829, -0.3707],
    innsbruck: [47.2692, 11.4041], vienne: [48.2082, 16.3738], vienna: [48.2082, 16.3738], wien: [48.2082, 16.3738],
    bale: [47.5596, 7.5886], basel: [47.5596, 7.5886], zurich: [47.3769, 8.5417], geneve: [46.2044, 6.1432],
    geneva: [46.2044, 6.1432], lausanne: [46.5197, 6.6323],
    copenhagen: [55.6761, 12.5683], copenhague: [55.6761, 12.5683], kobenhavn: [55.6761, 12.5683],
    amsterdam: [52.3676, 4.9041], delft: [52.0116, 4.3571], leiden: [52.1601, 4.4970], eindhoven: [51.4416, 5.4697],
    'esch-sur-alzette': [49.4958, 5.9806], luxembourg: [49.6116, 6.1319],
    munich: [48.1351, 11.5820], munchen: [48.1351, 11.5820], garching: [48.2489, 11.6532], berlin: [52.5200, 13.4050],
    stuttgart: [48.7758, 9.1829], heidelberg: [49.3988, 8.6724], mainz: [49.9929, 8.2473], hamburg: [53.5511, 9.9937],
    hambourg: [53.5511, 9.9937], jena: [50.9271, 11.5892], karlsruhe: [49.0069, 8.4037], erlangen: [49.5897, 11.0078],
    london: [51.5074, -0.1278], londres: [51.5074, -0.1278], oxford: [51.7520, -1.2577], cambridge: [52.2053, 0.1218],
    madrid: [40.4168, -3.7038], barcelona: [41.3874, 2.1686], barcelone: [41.3874, 2.1686],
    milan: [45.4642, 9.1900], rome: [41.9028, 12.4964], florence: [43.7696, 11.2558],
    bruxelles: [50.8503, 4.3517], brussels: [50.8503, 4.3517], louvain: [50.8798, 4.7005], leuven: [50.8798, 4.7005],
    stockholm: [59.3293, 18.0686], gothenburg: [57.7089, 11.9746], helsinki: [60.1699, 24.9384], oslo: [59.9139, 10.7522],
    prague: [50.0755, 14.4378], varsovie: [52.2297, 21.0122], warsaw: [52.2297, 21.0122], dublin: [53.3498, -6.2603]
};

function normalizeKey(text) {
    return String(text || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();
}

function clean(value) {
    if (value === null || value === undefined) return '';
    return String(value).replace(/\s+/g, ' ').trim();
}

// "Esch-sur-Alzette - Luxembourg" -> { city: "Esch-sur-Alzette", country: "Luxembourg" }
// Le séparateur est un tiret entouré d'au moins un espace, ou une virgule.
function parseCityCountry(text) {
    const parts = clean(text).split(/\s+-\s*|\s*-\s+|,/).map(s => s.trim()).filter(Boolean);
    if (!parts.length) return null;
    const rawCountry = parts.length > 1 ? parts.pop() : '';
    const city = parts.join(' - ');
    const country = COUNTRY_NAMES[normalizeKey(rawCountry)] || rawCountry;
    return { city, country };
}

function jobType(title) {
    if (!title) return null;
    const found = JOB_KEYWORDS.find(([, regex]) => regex.test(title));
    return found ? found[0] : 'other';
}

// Construit une étape (stage ou poste) ; null si la ligne ne contient rien
function buildEntry(fields) {
    const place = parseCityCountry(fields.place);
    const entry = { title: fields.title, institution: fields.institution };
    if (fields.supervisor) entry.supervisor = fields.supervisor;
    if (place) Object.assign(entry, place);
    return (entry.title || entry.institution || place) ? entry : null;
}

function parseSheet(rows, year) {
    const people = [];
    // Ligne 1 = titres des colonnes
    rows.slice(1).forEach(row => {
        const c = i => clean(row[i]);
        const lastName = c(0);
        const firstName = c(1);
        if (!lastName && !firstName) return;

        const internship = buildEntry({ title: c(2), institution: c(3), place: c(4) });
        const job = buildEntry({ title: c(5), institution: c(6), place: c(7), supervisor: c(8) });
        if (job) job.type = jobType(job.title) || 'other';

        people.push({ firstName, lastName, promotion: year, internship, job });
    });
    return people;
}

// Lit le fichier Excel et renvoie { promotions, people }
async function loadPromotionsFromXlsx() {
    const response = await fetch(XLSX_FILE, { cache: 'no-cache' });
    if (!response.ok) throw new Error(`${XLSX_FILE} introuvable (${response.status})`);
    const workbook = XLSX.read(await response.arrayBuffer(), { type: 'array' });

    let people = [];
    const years = [];
    workbook.SheetNames.forEach(sheetName => {
        const match = sheetName.match(/(\d{4})/);
        if (!match) return;
        const year = match[1];
        const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: null });
        years.push(year);
        people = people.concat(parseSheet(rows, year));
    });

    const promotions = {};
    years.sort((a, b) => b - a).forEach((year, i) => {
        promotions[year] = {
            name: PROMOTION_NAMES[year] || `Promotion ${year}`,
            color: PROMOTION_PALETTE[i % PROMOTION_PALETTE.length]
        };
    });

    people.forEach((person, i) => { person.id = i; });
    return { promotions, people };
}

// ---------- Coordonnées des villes ----------
const GEOCODE_CACHE_KEY = 'qtep-geocode-v1';

function readGeocodeCache() {
    try {
        return JSON.parse(localStorage.getItem(GEOCODE_CACHE_KEY)) || {};
    } catch (e) {
        return {};
    }
}

function writeGeocodeCache(cache) {
    try {
        localStorage.setItem(GEOCODE_CACHE_KEY, JSON.stringify(cache));
    } catch (e) { /* stockage indisponible : on recherchera à nouveau la prochaine fois */ }
}

function entriesOf(people) {
    return people.flatMap(p => [p.internship, p.job]).filter(e => e && e.city);
}

// Place les villes connues immédiatement ; renvoie la liste des villes à chercher en ligne
function applyKnownCoordinates(people) {
    const cache = readGeocodeCache();
    const unknown = new Map();
    entriesOf(people).forEach(entry => {
        const key = normalizeKey(entry.city);
        const cacheKey = `${key}|${normalizeKey(entry.country)}`;
        const coords = KNOWN_CITIES[key] || cache[cacheKey];
        if (coords) {
            [entry.lat, entry.lng] = coords;
        } else {
            unknown.set(cacheKey, { city: entry.city, country: entry.country });
        }
    });
    return unknown;
}

// Cherche les villes inconnues sur OpenStreetMap (1 requête par seconde maximum)
async function geocodeUnknownCities(unknown) {
    const cache = readGeocodeCache();
    for (const [cacheKey, { city, country }] of unknown) {
        try {
            const query = encodeURIComponent([city, country].filter(Boolean).join(', '));
            const res = await fetch(`https://nominatim.openstreetmap.org/search?format=json&limit=1&q=${query}`);
            const results = await res.json();
            if (results.length) {
                cache[cacheKey] = [Number(results[0].lat), Number(results[0].lon)];
                writeGeocodeCache(cache);
            }
        } catch (e) {
            console.warn('Ville introuvable :', city, country, e);
        }
        await new Promise(resolve => setTimeout(resolve, 1100));
    }
}
