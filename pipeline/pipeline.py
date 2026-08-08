# -*- coding: utf-8 -*-
import json, os, re, unicodedata, math
import openpyxl
from collections import Counter

HERE = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(os.path.dirname(HERE), 'Books TBR.xlsx')
wb = openpyxl.load_workbook(XLSX, read_only=True, data_only=True)
ws = wb['Library']
books = []
for row in ws.iter_rows(min_row=2, values_only=True):
    if not row[1] or not str(row[1]).strip(): continue
    typ = str(row[0]).strip() if row[0] else ''
    author = str(row[2]).strip() if row[2] and row[2] != '-' else ''
    pages = None
    if len(row) > 3 and row[3] not in (None, '', '-'):
        try: pages = int(float(str(row[3]).strip()))
        except ValueError: pass
    rating = None
    if len(row) > 4 and row[4] not in (None, '', '-'):
        try: rating = int(float(str(row[4]).strip()))
        except ValueError: pass
    status = str(row[5]).strip() if len(row) > 5 and row[5] else ''
    rec = {'title': str(row[1]).strip(), 'source': typ, 'author': author}
    if pages: rec['pages'] = pages
    if rating: rec['my_rating'] = rating
    if status == 'read': rec['gr_shelf'] = 'read'
    books.append(rec)

SCRATCH = HERE
def strip_diacritics(s):
    return ''.join(c for c in unicodedata.normalize('NFKD', s) if not unicodedata.combining(c))

def norm(s):
    s = strip_diacritics(s or '').lower()
    s = re.sub(r'[^a-z0-9 ]', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s

EXT_RE = re.compile(r'\.(epub|pdf|mobi|odt|azw3|txt|PDF)\b', re.I)

def tidy_author(cand):
    """Normalise an author string: strip junk, reorder 'Last, First' -> 'First Last'."""
    if not cand: return ''
    cand = re.sub(r'\[.*?\]', '', cand)          # drop [Last, First] duplicates
    cand = cand.replace('_', '').strip(' -.,;')
    if re.search(r'\d{3,}', cand) or len(cand) > 45:
        return ''
    # take first author if a list, but keep "First Last" whole
    if ';' in cand:
        cand = cand.split(';')[0].strip()
    # reorder "Last, First"
    if ',' in cand and '&' not in cand:
        a, b = cand.split(',', 1)
        a, b = a.strip(), b.strip()
        if a and b and ' ' not in a and len(b) < 25:
            cand = f'{b} {a}'
        else:
            cand = a
    return cand.strip(' -.,;')

def clean_title(raw, author_hint):
    t = raw.strip()
    author = tidy_author(author_hint or '')
    # strip file extension up front
    t = EXT_RE.sub('', t)
    # Anna's Archive pattern:  Title -- Author -- publisher -- isbn -- hash -- Anna's Archive
    if '--' in t and ('Anna' in t or 'Archive' in t or re.search(r'--\s*[0-9a-f]{16,}', t) or 'isbn' in t.lower()):
        parts = [p.strip() for p in t.split('--')]
        t = parts[0]
        if not author and len(parts) > 1:
            author = tidy_author(parts[1])
    # z-library pattern:  Title (Author) (z-library ...)  OR  Title (z-lib.org)
    if 'z-lib' in t.lower() or 'z-library' in t.lower() or '1lib' in t.lower():
        t = re.sub(r'\((?:z-lib|z-library|1lib)[^)]*\)', '', t, flags=re.I)
    t = re.sub(r'\(z-?library\)', '', t, flags=re.I)
    # strip trailing "-- ..." remnants / lone hashes
    t = re.sub(r'--.*$', '', t).strip()
    # a trailing "(Author Name)" -> extract to author, strip from title
    m = re.search(r'\(([^()]+)\)\s*$', t)
    if m:
        cand = m.group(1).strip()
        looks_person = (' ' in cand or ',' in cand) and not re.search(r'\d', cand) \
                       and not re.search(r'(?i)\b(etc|edition|series|vol|book|university|center|classic|the )\b', cand) \
                       and len(cand) < 40
        if looks_person:
            if not author:
                author = tidy_author(cand)
            t = t[:m.start()].strip()
    # drop a dangling unmatched "(" fragment  e.g. "Meaning in Life (The University Center"
    if t.count('(') > t.count(')'):
        t = t[:t.rfind('(')].strip()
    # strip trailing year in parens or bare trailing 19xx/20xx
    t = re.sub(r'\(\d{4}\)\s*$', '', t).strip()
    t = re.sub(r'\b(19|20)\d{2}\b\s*$', '', t).strip()
    # strip trailing " by Author" (only if no author yet, and >=2 words before "by")
    if not author:
        m = re.search(r'\bby\s+([A-Z][A-Za-z.\-]+(?:\s+[A-Z][A-Za-z.\-]+){0,3})\s*$', t)
        if m and len(t[:m.start()].split()) >= 2:
            author = m.group(1).strip()
            t = t[:m.start()].strip()
    else:
        t = re.sub(r'\s+by\s+[A-Z][A-Za-z.\-]+(?:\s+[A-Z][A-Za-z.\-]+){0,3}\s*$', '', t).strip()
    # tidy whitespace / separators
    t = t.replace('_', ' ')
    t = re.sub(r'\s+', ' ', t).strip(' -.,')
    if not t:
        t = re.sub(r'\.(epub|pdf|mobi|odt)$','',raw.strip(), flags=re.I)
    return t, author.strip(' -.,;')

for b in books:
    ct, ca = clean_title(b['title'], b['author'])
    b['title'] = ct
    b['author'] = ca

# ---------------------------------------------------------------------------
# 2. CLASSIFICATION DATA
# ---------------------------------------------------------------------------
# Genres used across the map
# Manga, Comics, Literary Fiction, Classics, Fantasy, Science Fiction,
# Mystery & Thriller, Horror, Romance, Young Adult, Children's,
# Poetry & Drama, Science & Nature, Mathematics, Philosophy,
# Psychology & Self-Help, Business & Finance, History, Biography & Memoir,
# Spirituality & Religion, Politics & Society, Arts & Film,
# Health & Fitness, Reference & Learning

MANGA = {  # normalized title -> True
 'elfen lied','jojo s bizarre adventure','psyren','20th 21st century boys','death note',
 'monster','bleach','made in abyss','demon slayer','fullmetal alchemist','hellsing','tokyo ghoul',
 'jujutsu kaisen','orange','dragon ball super','dragon ball z redownload','dragon ball z',
 'girl s last tour','erased','annarasumanara','planetes','record of ragnarok','god s lie','ping pong',
 'all you need is kill','i am a hero','i am hero','inuyashiki','one punch man','saint onisan','saint oniisan',
 'blue spring','chainsaw man','chainsawman','black butler','fire punch','akira','nana','sun ken rock',
 'blissful land','kaichou wa maid sama','golden kamuy','hunter x hunter','kaguya sama love is war',
 'dorohedoro','akumetsu','claymore','3 gatsu no lion','deadman wonderland','the way of the househusband',
 'high rise invasion','the girl from the other side','solanin','my hero academia','the laundromat woman',
 'kaiju no 8','the walking man','uzumaki','vagabound','vagabond','spy x family',
 'the way of househusband','death note special one shot','deadman wonderland',
}
MANGA_AUTHORS = {'boichi','satoru noda','yoshihiro togashi','aka akasaka','q hayashida','hiro fujiwara',
 'ichimon izumi','yoshiaki tabata','norihiro yagi','junji ito','mushashi miyamoto'}

COMIC_KEYS = ['batman','superman','justice league','deadpool','doctor strange','green lantern','suicide squad',
 'teen titans','deathstroke','flashpoint','flash rebirth','crisis on infinite earths','darkseid','joker',
 'sin city','sandman','watchmen','v for vendetta','the walking dead','transformers','g i joe','american vampire',
 'fables','wicked the divine','house of penance','stitched','the last of us','wraith','earth 2','trinity',
 'supergirl','two face','detectives comics','dark nights','dark days','injustice','all star superman',
 'reign of doomsday','our world at war','52 aftermath','gotham academy','star wars','star trek',
 'ramayana 3392','avatarex','sadhu','maus','persepolis','graphic novel','graphic adaptation',
 'anne frank s diary','wonderwoman','doomsday']
COMIC_TITLES = {'opus','seconds','the complete persepolis','maus','watchmen','fables','wraith','stitched',
 'everything is teeth','parasite graphic novel storyboards','i think therefore i draw'}

# Author -> genre (normalized author substring match). Order matters loosely; specific first.
AUTHOR_GENRE = {
 # Fantasy
 'brandon sanderson':'Fantasy','scott lynch':'Fantasy','joe abercrombie':'Fantasy','j r r tolkien':'Fantasy',
 'tolkien':'Fantasy','george r r martin':'Fantasy','neil gaiman':'Fantasy','v e schwab':'Fantasy',
 'leigh bardugo':'Fantasy','cassandra clare':'Fantasy','philip pullman':'Fantasy','christopher paolini':'Fantasy',
 'sarah j maas':'Fantasy','susanna clarke':'Fantasy','john gwynne':'Fantasy','chris wooding':'Fantasy',
 'erin morgenstern':'Fantasy','katherine arden':'Fantasy','madeline miller':'Fantasy','terry pratchett':'Fantasy',
 'amish tripathi':'Fantasy','ananda neelakantan':'Fantasy','christopher buehlman':'Fantasy','patricia a mckillip':'Fantasy',
 'kathryn lasky':'Fantasy','shannon messenger':'Children\'s','deborah harkness':'Fantasy','r f kuang':'Fantasy',
 'natasha pulley':'Fantasy','akshat gupta':'Fantasy','ashwin sanghi':'Fantasy','vineet bajpai':'Fantasy',
 'kevin misal':'Fantasy','stephenie meyer':'Fantasy','eoin colfer':'Children\'s',
 # Science Fiction
 'isaac asimov':'Science Fiction','arthur c clarke':'Science Fiction','kurt vonnegut':'Science Fiction',
 'douglas adams':'Science Fiction','neal stephenson':'Science Fiction','blake crouch':'Science Fiction',
 'peter watts':'Science Fiction','stephen baxter':'Science Fiction','octavia butler':'Science Fiction',
 'samantha harvey':'Science Fiction','mainak dhar':'Science Fiction','scott westerfeld':'Young Adult',
 'james dashner':'Young Adult','veronica roth':'Young Adult','suzanne collins':'Young Adult',
 # Mystery & Thriller
 'agatha christie':'Mystery & Thriller','arthur conan doyle':'Mystery & Thriller','raymond chandler':'Mystery & Thriller',
 'dashiell hammett':'Mystery & Thriller','dan brown':'Mystery & Thriller','david baldacci':'Mystery & Thriller',
 'robert ludlum':'Mystery & Thriller','gillian flynn':'Mystery & Thriller','lee child':'Mystery & Thriller',
 'robert galbraith':'Mystery & Thriller','alex michaelides':'Mystery & Thriller','jeffrey archer':'Mystery & Thriller',
 'karen m mcmanus':'Mystery & Thriller','g k chesterton':'Mystery & Thriller','michael palmer':'Mystery & Thriller',
 'kate mosse':'Mystery & Thriller','robert coram':'Biography & Memoir','vincent bugliosi':'Mystery & Thriller',
 'christopher berry dee':'Mystery & Thriller',
 # Horror
 'clive barker':'Horror','h p lovecraft':'Horror','stephen king':'Horror','r l stine':'Children\'s',
 'joe hill':'Horror','stephen graham jones':'Horror','alvin schwartz':'Children\'s','scott smith':'Horror',
 # Romance
 'nicholas sparks':'Romance','colleen hoover':'Romance','jojo moyes':'Romance','sally thorne':'Romance',
 'rachael lippincott':'Romance','taylor jenkin reids':'Romance','preeti shenoy':'Romance','audrey niffenegger':'Romance',
 # Young Adult
 'john green':'Young Adult','stephen chbosky':'Young Adult','jay asher':'Young Adult','angie thomas':'Young Adult',
 'maureen johnson':'Young Adult','benjamin alire saenz':'Young Adult','anna marie mclemore':'Young Adult',
 'kerstin gier':'Young Adult','k a applegate':'Young Adult','lord emery':'Young Adult','nichola yoong':'Young Adult',
 # Children's
 'roald dahl':'Children\'s','jeff kinney':'Children\'s','enid blyton':'Children\'s','francesca simon':'Children\'s',
 'lemony snicket':'Children\'s','tove jansson':'Children\'s','elizabeth goudge':'Children\'s','john bellairs':'Children\'s',
 'deborah howe':'Children\'s','laura ingalls wilder':'Children\'s','virginia sorensen':'Children\'s','anita ganeri':'Children\'s',
 'terry deary':'Children\'s','kinney':'Children\'s','antoine de saint exupery':'Children\'s','lewis carroll':'Children\'s',
 'richard adams':'Children\'s','hans christian andersen':'Children\'s','anthony horowitz':'Young Adult',
 # Poetry & Drama
 'emily dickinson':'Poetry & Drama','pablo neruda':'Poetry & Drama','sylvia plath':'Poetry & Drama',
 'walt whitman':'Poetry & Drama','sarah kay':'Poetry & Drama','rainer':'Poetry & Drama','kahlil gibran':'Poetry & Drama',
 'homer':'Poetry & Drama','dante':'Poetry & Drama','john milton':'Poetry & Drama','geoffrey chaucer':'Poetry & Drama',
 'javed akhtar':'Poetry & Drama','rahat indori':'Poetry & Drama','gopaldas neeraj':'Poetry & Drama',
 'kumar vishwas':'Poetry & Drama','faiz ahmed faiz':'Poetry & Drama','harivanshrai bachan':'Poetry & Drama',
 'shiv kumar batalvi':'Poetry & Drama','kabir':'Poetry & Drama','neil simon':'Poetry & Drama','anne carson':'Poetry & Drama',
 'tracy k smith':'Poetry & Drama','porter max':'Poetry & Drama','sarah kay':'Poetry & Drama','jeremy o harris':'Poetry & Drama',
 'wendy cope':'Poetry & Drama','t s eliot':'Poetry & Drama','george bernard shaw':'Poetry & Drama',
 'kalidasa':'Poetry & Drama','mishra piyush':'Poetry & Drama','ashutosh rana':'Poetry & Drama','sarojini naidu':'Poetry & Drama',
 'emily bronte':'Classics','leaves of grass':'Poetry & Drama',
 # Literary Fiction (modern / contemporary)
 'haruki murakami':'Literary Fiction','sally rooney':'Literary Fiction','ottessa moshfegh':'Literary Fiction',
 'donna tartt':'Literary Fiction','zadie smith':'Literary Fiction','orhan pamuk':'Literary Fiction',
 'kazuo ishiguro':'Literary Fiction','cormac mccarthy':'Literary Fiction','toni morrison':'Literary Fiction',
 'jhumpa lahiri':'Literary Fiction','aravind adiga':'Literary Fiction','amitav ghosh':'Literary Fiction',
 'arundhati roy':'Literary Fiction','michael chabon':'Literary Fiction','margaret atwood':'Literary Fiction',
 'ocean voung':'Literary Fiction','ian mcewan':'Literary Fiction','matt haig':'Literary Fiction',
 'maggie o farrell':'Literary Fiction','david mitchell':'Literary Fiction','naoise dolan':'Literary Fiction',
 'yukio mishima':'Literary Fiction','osamu dazai':'Literary Fiction','ryunosuke akutagawa':'Literary Fiction',
 'akutagawa':'Literary Fiction','yoko ogawa':'Literary Fiction','yoko tawada':'Literary Fiction',
 'jorge luis borges':'Literary Fiction','borges':'Literary Fiction','roberto bolano':'Literary Fiction',
 'lydia davis':'Literary Fiction','amy hempel':'Literary Fiction','flannery o connor':'Literary Fiction',
 'alice munro':'Literary Fiction','claire keegan':'Literary Fiction','denis johnson':'Literary Fiction',
 'nathanael west':'Literary Fiction','john williams':'Literary Fiction','paulo coelho':'Literary Fiction',
 'khaled hosseini':'Literary Fiction','min jin lee':'Literary Fiction','chitra banerjee divakaruni':'Literary Fiction',
 'r k narayan':'Literary Fiction','rk narayan':'Literary Fiction','ruskin bond':'Literary Fiction',
 'anita desai':'Literary Fiction','kiran desai':'Literary Fiction','vikram seth':'Literary Fiction',
 'khushwant singh':'Literary Fiction','chetan bhagat':'Literary Fiction','anuja chauhan':'Literary Fiction',
 'sudha murty':'Literary Fiction','sudha murthy':'Literary Fiction','gregory d roberts':'Literary Fiction',
 'yann martel':'Literary Fiction','markus zusak':'Literary Fiction','fredrik backman':'Literary Fiction',
 'anthony doerr':'Literary Fiction','celeste ng':'Literary Fiction','maggie shipstead':'Literary Fiction',
 'patricia lockwood':'Literary Fiction','lucy ellmann':'Literary Fiction','m l rio':'Literary Fiction',
 'r f kuang babel':'Literary Fiction','hank green':'Literary Fiction','john scalzi':'Science Fiction',
 'saou ichikawa':'Literary Fiction','marjan kamali':'Literary Fiction','amor towles':'Literary Fiction',
 'ngozi adichie':'Literary Fiction','half of a yellow sun':'Literary Fiction','junichiro tanizaki':'Literary Fiction',
 'tanizaki':'Literary Fiction','satoshi yagisawa':'Literary Fiction','toshikazu kawaguchi':'Literary Fiction',
 'gavin extence':'Literary Fiction','saramago':'Literary Fiction','george saunders':'Literary Fiction',
 'mark haddon':'Literary Fiction','joan didion':'Literary Fiction','david sedaris':'Literary Fiction',
 'nidhi upadhyay':'Literary Fiction','shivalik bakshi':'Literary Fiction','samsara saksham garg':'Literary Fiction',
 'chowringhee':'Literary Fiction','mani shankar':'Literary Fiction','mordecai richler':'Literary Fiction',
 'walter tevis':'Literary Fiction','david benioff':'Literary Fiction','michael thomas ford':'Young Adult',
 'daisy jones':'Literary Fiction','qiu miaojin':'Literary Fiction','hayley phelan':'Literary Fiction',
 'silvina ocampo':'Literary Fiction','carrington leonora':'Literary Fiction','fernanda melchor':'Literary Fiction',
 'tatiana de rosnay':'Literary Fiction','joe sumner':'Comics & Graphic Novels',
 # Classics (older canon)
 'leo tolstoy':'Classics','fyodor dostoevsky':'Classics','fyodor dostoyevsky':'Classics','charles dickens':'Classics',
 'jane austen':'Classics','herman melville':'Classics','mark twain':'Classics','nathaniel hawthorne':'Classics',
 'joseph conrad':'Classics','franz kafka':'Classics','nikolai gogol':'Classics','nikolay gogol':'Classics',
 'anton chekhov':'Classics','ernest hemingway':'Classics','f scott fitzgerald':'Classics','william faulkner':'Classics',
 'john steinbeck':'Classics','john steinback':'Classics','virginia woolf':'Classics','george eliot':'Classics',
 'gabriel garcia marquez':'Classics','gabriel garcia marquez':'Classics','vladimir nabokov':'Classics',
 'somerset maugham':'Classics','graham greene':'Classics','chinua achebe':'Classics','george orwell':'Classics',
 'aldous huxley':'Classics','william golding':'Classics','miguel de cervantes':'Classics','alexandre dumas':'Classics',
 'victor hugo':'Classics','mary shelley':'Classics','bram stoker':'Classics','oscar wilde':'Classics',
 'daphne du maurier':'Classics','robert penn warren':'Classics','theodore dreiser':'Classics','sinclair lewis':'Classics',
 'sinclair upton':'Classics','upton sinclair':'Classics','ralph ellison':'Classics','richard wright':'Classics',
 'zora neale hurston':'Classics','harriet beecher stowe':'Classics','stowe':'Classics','robert musil':'Classics',
 'mikhail bulgakov':'Classics','robert graves':'Classics','stendhal':'Classics','elizabeth gaskell':'Classics',
 'thomas hardy':'Classics','henry james':'Classics','james baldwin':'Classics','margaret mitchell':'Classics',
 'ken kesey':'Classics','jerzy kosinski':'Classics','walker percy':'Classics','john fante':'Classics',
 'charles bukowski':'Classics','p g wodehouse':'Classics','wodehouse':'Classics','simone de beauvoir':'Classics',
 'jean rhys':'Classics','boethius':'Philosophy','sophie s world':'Philosophy','emily bronte':'Classics',
 'charlotte bronte':'Classics','leo tolstoy':'Classics','solzhenitsyn':'Classics','solzhenitsyn':'Classics',
 'alexandr solzhenitsyn':'Classics','robert louis stevenson':'Classics','r l stevenson':'Classics',
 'nikos kazantzakis':'Classics','premchand':'Classics','sharatchandar chattopadhyay':'Classics',
 'jay shankar prasad':'Classics','manto':'Classics','saadat hasan manto':'Classics','sinclair':'Classics',
 'rabindranath tagore':'Classics','bamkinchandra':'Classics','arthur golden':'Classics','anne frank':'Biography & Memoir',
 'joyce carey':'Classics','joyce cary':'Classics','leonora carrington':'Literary Fiction','henri barbusse':'Classics',
 'leonid andreyev':'Classics','sarah bakewell':'Philosophy','budd schulberg':'Classics','sinclair upton':'Classics',
 'kahlil':'Poetry & Drama','vinod kumar shukl':'Literary Fiction','march william':'Classics','tobias wolff':'Biography & Memoir',
 # Science & Nature
 'richard dawkins':'Science & Nature','bill bryson':'Science & Nature','bryson bill':'Science & Nature',
 'brian greene':'Science & Nature','stephen hawking':'Science & Nature','stephen w hawking':'Science & Nature',
 'michio kaku':'Science & Nature','vaclav smil':'Science & Nature','neil shubin':'Science & Nature',
 'merlin sheldrake':'Science & Nature','charles darwin':'Science & Nature','albert einstein':'Science & Nature',
 'roger penrose':'Science & Nature','steven pinker':'Science & Nature','matthew walker':'Science & Nature',
 'david eagleman':'Science & Nature','bernd heinrich':'Science & Nature','peter wohlleben':'Science & Nature',
 'bernard wood':'Science & Nature','john m barry':'Science & Nature','richard feynman':'Science & Nature',
 'richard p feynman':'Science & Nature','james gleick':'Science & Nature','carl sagan':'Science & Nature',
 'lisa a urry':'Science & Nature','kevin langford':'Science & Nature','a b bhattacharya':'Science & Nature',
 'mark miodownik':'Science & Nature','ed yong':'Science & Nature','david deutsch':'Science & Nature',
 'robert m sapolsky':'Science & Nature','robert sapolsky':'Science & Nature','frances e jensen':'Science & Nature',
 'daniel j siegel':'Science & Nature','kelly and zach weinersmith':'Science & Nature','safi bahcall':'Science & Nature',
 'stephanie mcmurrich':'Health & Fitness','holly jean buck':'Science & Nature','david wallace wells':'Science & Nature',
 'david stainforth':'Science & Nature','matt parker':'Mathematics','alex bellos':'Mathematics','graham everest':'Mathematics',
 'clifford a pickover':'Mathematics','cliffordd a pickover':'Mathematics','tim harford':'Mathematics',
 # Mathematics handled above; Health & Fitness
 'mark rippetoe':'Health & Fitness','john little':'Health & Fitness','scott douglas':'Health & Fitness',
 'leslie kaminoff':'Health & Fitness','greg douchette':'Health & Fitness','ross edgley':'Health & Fitness',
 'alex hutchinson':'Health & Fitness','emily nagoski':'Health & Fitness','ian kerner':'Health & Fitness',
 'doug abrams':'Health & Fitness','harville hendrix':'Psychology & Self-Help','david epstein':'Psychology & Self-Help',
 'olli sovijarvi':'Health & Fitness',
 # Philosophy
 'friedrich nietzsche':'Philosophy','arthur schopenhauer':'Philosophy','soren kierkegaard':'Philosophy',
 'jean paul sartre':'Philosophy','albert camus':'Philosophy','plato':'Philosophy','aristotle':'Philosophy',
 'michel de montaigne':'Philosophy','heraclitus':'Philosophy','epicurus':'Philosophy','seneca':'Philosophy',
 'marcus aurelius':'Philosophy','bertrand russel':'Philosophy','bertrand russell':'Philosophy','russel':'Philosophy',
 'alain de bottom':'Philosophy','alain de botton':'Philosophy','simone weil':'Philosophy','eugene thacker':'Philosophy',
 'nolen gertz':'Philosophy','todd may':'Philosophy','susan wolf':'Philosophy','mark fisher':'Philosophy',
 'jostein gaarder':'Philosophy','sir thomas more':'Philosophy','voltare':'Philosophy','voltaire':'Philosophy',
 'jordan b peterson':'Psychology & Self-Help','dr jordan b peterson':'Psychology & Self-Help','r d laing':'Philosophy',
 'ernest becker':'Philosophy','will durant':'Philosophy','noam chomsky':'Politics & Society','steven connor':'Philosophy',
 'publius syrus':'Philosophy','john graham':'Philosophy','john j kaag':'Philosophy','ralph waldo emerson':'Philosophy',
 'henry david thoreau':'Philosophy','thoreau':'Philosophy','daniel klein':'Philosophy','robin hanson':'Philosophy',
 'karl marx':'Politics & Society','sam harris':'Philosophy','anicius boethius':'Philosophy',
 # Psychology & Self-Help / Productivity
 'ryan holiday':'Psychology & Self-Help','james clear':'Psychology & Self-Help','cal newport':'Psychology & Self-Help',
 'steven pressfield':'Psychology & Self-Help','daniel kahneman':'Psychology & Self-Help','jonathan haidt':'Psychology & Self-Help',
 'robert wright':'Psychology & Self-Help','eckhart tolle':'Spirituality & Religion','benjamin hardy':'Psychology & Self-Help',
 'dale carnegie':'Psychology & Self-Help','ichiro kishimi':'Psychology & Self-Help','robert green':'Psychology & Self-Help',
 'robert greene':'Psychology & Self-Help','steven kotler':'Psychology & Self-Help','jenny blake':'Psychology & Self-Help',
 'chris bailey':'Psychology & Self-Help','sonke ahrens':'Psychology & Self-Help','james carse':'Psychology & Self-Help',
 'greg mckeown':'Psychology & Self-Help','ali abdaal':'Psychology & Self-Help','austin kleon':'Psychology & Self-Help',
 'mark manson':'Psychology & Self-Help','david allen':'Psychology & Self-Help','robert cialdini':'Psychology & Self-Help',
 'cialdini':'Psychology & Self-Help','malcolm gladwell':'Psychology & Self-Help','angela duckworth':'Psychology & Self-Help',
 'susan cain':'Psychology & Self-Help','maria konnikova':'Psychology & Self-Help','philip e tetlock':'Psychology & Self-Help',
 'oliver burkeman':'Psychology & Self-Help','gloria mark':'Psychology & Self-Help','robin m hogarth':'Psychology & Self-Help',
 'todd rose':'Psychology & Self-Help','nir eyal':'Psychology & Self-Help','charles duhigg':'Psychology & Self-Help',
 'daniel gilbert':'Psychology & Self-Help','jon ronson':'Psychology & Self-Help','thomas erikson':'Psychology & Self-Help',
 'stephen r covey':'Psychology & Self-Help','tony robbins':'Psychology & Self-Help','logan ury':'Psychology & Self-Help',
 'lori gottlieb':'Psychology & Self-Help','bessel van der kolk':'Psychology & Self-Help','carl gustav jung':'Psychology & Self-Help',
 'carl jung':'Psychology & Self-Help','c g jung':'Psychology & Self-Help','sigmund freud':'Psychology & Self-Help',
 'shwetabh gangwar':'Psychology & Self-Help','james allen':'Psychology & Self-Help','jocko willink':'Psychology & Self-Help',
 'thom hartmann':'Psychology & Self-Help','david mcraney':'Psychology & Self-Help','richard wiseman':'Psychology & Self-Help',
 'randy j peterson':'Psychology & Self-Help','harry styles':'Psychology & Self-Help','patrick king':'Psychology & Self-Help',
 'karp harvey':'Psychology & Self-Help','jane nelsen':'Psychology & Self-Help','keri smith':'Psychology & Self-Help',
 'julia cameron':'Psychology & Self-Help','ryan holiday':'Psychology & Self-Help','richard koch':'Psychology & Self-Help',
 'gloria mark':'Psychology & Self-Help','scott barry':'Psychology & Self-Help','robin sharma':'Psychology & Self-Help',
 'hank green an absolutely':'Science Fiction','harville hendrix':'Psychology & Self-Help','doug abrams':'Health & Fitness',
 # Business & Finance
 'ray dalio':'Business & Finance','nassim nicholas taleb':'Business & Finance','nicholas nassim taleb':'Business & Finance',
 'benjamin graham':'Business & Finance','peter lynch':'Business & Finance','peter s lynch':'Business & Finance',
 'jim collins':'Business & Finance','clayton m christensen':'Business & Finance','michael lewis':'Business & Finance',
 'thomas piketty':'Business & Finance','liaquat ahamed':'Business & Finance','charles d ellis':'Business & Finance',
 'adam lashinsky':'Business & Finance','morgan housel':'Business & Finance','jim rogers':'Business & Finance',
 'phil knight':'Business & Finance','peter thiel':'Business & Finance','richard branson':'Business & Finance',
 'walter isaacson':'Biography & Memoir','ashlee vance':'Biography & Memoir','michael dell':'Business & Finance',
 'alice schroeder':'Business & Finance','saurabh mukherjea':'Business & Finance','russell brunson':'Business & Finance',
 'chris guillebeau':'Business & Finance','byron sharp':'Business & Finance','david ogilvy':'Business & Finance',
 'eugene m schwartz':'Business & Finance','michael porter':'Business & Finance','michael masterson':'Business & Finance',
 'daniel priestley':'Business & Finance','noah kagan':'Business & Finance','gino wickman':'Business & Finance',
 'reed hastings':'Business & Finance','duncan clark':'Business & Finance','david s rose':'Business & Finance',
 'andrew chen':'Business & Finance','geoffrey a moore':'Business & Finance','ben horowitz':'Business & Finance',
 'eric ries':'Business & Finance','safi bahcall':'Business & Finance','jill schlesinger':'Business & Finance',
 'thornton oglove':'Business & Finance','joel greenblatt':'Business & Finance','greenblatt':'Business & Finance',
 'aswath damodaran':'Business & Finance','howard marks':'Business & Finance','steve nison':'Business & Finance',
 'thomas sowell':'Business & Finance','henry hazlitt':'Business & Finance','r vaidyanathan':'Business & Finance',
 'anita raghvan':'Business & Finance','binod chaudhary':'Business & Finance','sarthak ahuja':'Business & Finance',
 'jill schlesinger':'Business & Finance','richard koch':'Business & Finance','w chan kim':'Business & Finance',
 'adam smith':'Business & Finance','warren buffett':'Business & Finance','colin bryar':'Business & Finance',
 'karen berman':'Business & Finance','tony robbins':'Psychology & Self-Help','chip heath':'Business & Finance',
 'kate welling':'Business & Finance','carey and morris':'Business & Finance','russell napier':'Business & Finance',
 'david g thomson':'Business & Finance','herminia ibarra':'Business & Finance','paul millerd':'Business & Finance',
 'mario gabelli':'Business & Finance','hamish mcdonald':'Business & Finance','james crabtree':'Business & Finance',
 'saurabh':'Business & Finance','n r narayana':'Business & Finance','simon anholt':'Business & Finance',
 # History
 'william dalrymple':'History','yuval noah harari':'History','doris kearns goodwin':'History','robert a caro':'History',
 'donald kagan':'History','arnold toynbee':'History','d c somervell':'History','eric hobsbawm':'History',
 'abraham eraly':'History','irfan habib':'History','satish chandra':'History','upinder singh':'History',
 'arjun dev':'History','b v rao':'History','peter frankopan':'History','charles allen':'History','ashoka':'History',
 'simon sebag montefiore':'History','stalin':'History','anthony beevor':'History','jill lepore':'History',
 'peter conradi':'History','roy moxham':'History','ratna ghosh':'History','manoshi sinha':'History','rima hooja':'History',
 'ranjit desai':'History','patil vishwas':'History','vishwas patil':'History','shivaji sawant':'History',
 'kuldip nayar':'History','harman':'History','captivating history':'History','vasant kumar bawa':'History',
 's c ray chaudhary':'History','narain':'History','rajiv ahir':'History','veer savarkar':'History',
 'khilnani sunil':'History','sunil khilnani':'History','p sainath':'Politics & Society','a rahnema':'History',
 'ali rahnema':'History','arash azizi':'History','erik larson':'History','alfred lansing':'History',
 'dee brown':'History','toby wilkinson':'History','bibek debroy':'Spirituality & Religion','devdutt':'Spirituality & Religion',
 'satyarth nayak':'Spirituality & Religion','ami ganatra':'Spirituality & Religion','maharana pratap':'History',
 'lt gn kuldeep singh braat':'History','pavan k varma':'Spirituality & Religion','maharana':'History',
 'winston churchill':'History','encyclopedia of indian history':'History','anuja chandramouli':'Fantasy',
 'harinder singh sikka':'History','s hussain zaidi':'History',
 # Biography & Memoir
 'malcolm x':'Biography & Memoir','martin luther king':'Biography & Memoir','frank mccourt':'Biography & Memoir',
 'elton john':'Biography & Memoir','frederick douglass':'Biography & Memoir','lucy grealy':'Biography & Memoir',
 'elia kazan':'Biography & Memoir','john nathan':'Biography & Memoir','mike tyson':'Biography & Memoir',
 'cus d amato':'Biography & Memoir','david nasaw':'Biography & Memoir','andrew grove':'Biography & Memoir',
 'rafael nadal':'Biography & Memoir','robert kanigel':'Biography & Memoir','steve levy':'Biography & Memoir',
 'sebastian junger':'Biography & Memoir','michael finkel':'Biography & Memoir','tobias wolff':'Biography & Memoir',
 'vincent van gogh':'Arts & Film','vincent van gogh':'Arts & Film','helena merriman':'Biography & Memoir',
 'jean edward smith':'Biography & Memoir','rich cohen':'Biography & Memoir','alok kejriwal':'Biography & Memoir',
 'corrie ten boom':'Biography & Memoir','solomon northup':'Biography & Memoir','mary somerville':'Biography & Memoir',
 'in other words':'Biography & Memoir','mishra piyush':'Poetry & Drama','win mccormack':'Biography & Memoir',
 'iron ambition':'Biography & Memoir','henry david thoreau journals':'Biography & Memoir',
 # Spirituality & Religion
 'osho':'Spirituality & Religion','swami vivekananda':'Spirituality & Religion','paramahansa yogananda':'Spirituality & Religion',
 'a c bhaktivedanta':'Spirituality & Religion','sogyal rinpoche':'Spirituality & Religion','jed mckenna':'Spirituality & Religion',
 'jiddu krishnamurti':'Spirituality & Religion','krishnamurti':'Spirituality & Religion','om swami':'Spirituality & Religion',
 'g s chauhan':'Spirituality & Religion','anand sahib':'Spirituality & Religion','swami nityaswarupananda':'Spirituality & Religion',
 'j donald walters':'Spirituality & Religion','ranjit chaudhri':'Spirituality & Religion','melvin mcleod':'Spirituality & Religion',
 'wendy doniger':'Spirituality & Religion','visnu sarma':'Spirituality & Religion','sarah shaw':'Spirituality & Religion',
 'okakura':'Spirituality & Religion','sri sri paramahansa yogananda':'Spirituality & Religion','sorabh kudeshiya':'Spirituality & Religion',
 'kautilya':'History','vivek kumar':'Spirituality & Religion','neelotpal':'Poetry & Drama',
 # Politics & Society
 'noam chomsky':'Politics & Society','edward s herman':'Politics & Society','michael parenti':'Politics & Society',
 'william blum':'Politics & Society','douglas murray':'Politics & Society','neil postman':'Politics & Society',
 'amartya sen':'Politics & Society','arun kumar':'Politics & Society','bryan caplan':'Politics & Society',
 'james c scott':'Politics & Society','bhagat singh':'Politics & Society','naxalite':'Politics & Society',
 'prakash singh':'Politics & Society','frank barat':'Politics & Society','s jaishankar':'Politics & Society',
 'neerja chowdhury':'Politics & Society','theodore dalrymple':'Politics & Society','bryan stevenson':'Politics & Society',
 'jonathan haidt anxious':'Politics & Society','free speech':'Politics & Society','barrington moore':'Politics & Society',
 'emile durkheim':'Politics & Society','durkheim':'Politics & Society','bell hooks':'Politics & Society',
 'joy buolam':'Politics & Society','mukul deva':'Mystery & Thriller','adolf hitler':'Politics & Society',
 'sun tzu':'Politics & Society','robert green 48':'Psychology & Self-Help','de mesquita':'Politics & Society',
 # Arts & Film
 'syd field':'Arts & Film','robert mckee':'Arts & Film','sidney lumet':'Arts & Film','david mamet':'Arts & Film',
 'blake snyder':'Arts & Film','walter murch':'Arts & Film','francis ching':'Arts & Film','andrew loomis':'Arts & Film',
 'george b bridgman':'Arts & Film','betty edwards':'Arts & Film','e h gombrich':'Arts & Film','robert l herbert':'Arts & Film',
 'carlo pedretti':'Arts & Film','john truby':'Arts & Film','truby':'Arts & Film','john walsh':'Arts & Film',
 'blain brown':'Arts & Film','james hoffmann':'Arts & Film','jens muller':'Arts & Film','david gibson':'Arts & Film',
 'robert rodriguez':'Arts & Film','neil strauss':'Psychology & Self-Help','robert rodriguez':'Arts & Film',
 'save the cat':'Arts & Film','film directing':'Arts & Film',
}

# Title keyword -> genre  (checked on normalized title). Order = priority.
TITLE_KEYWORDS = [
 # exam / reference / learning
 (['cfa level','cfa 2024','schwesernotes'],'Reference & Learning'),
 (['princeton sat','sat 2021'],'Reference & Learning'),
 (['genki','beginning japanese','easy japanese','learn bash','learn git','linux command','docker in practice',
   '100 days of code','python','efficient linux','edito b1','easy french','practice makes perfect french',
   'french grammar','french 1a','french from wikibooks','french short stories','spanish short','aula internacional',
   'como agua para chocolate','beginning japanese','architectural graphics','best practices for equity research',
   'family wealth preservation','private wealth management','emerging markets handbook','geopolitical alpha',
   'modern money mechanics','music business','the art of seo','affiliate','dotcom secrets','traffic secrets',
   'expert secrets','side hustle','scorecard marketing','how brands grow','competitive strategy',
   'the world atlas of coffee','cinematography','film directing','save the cat','on directing film',
   'the foundations of screenwriting','the photographer','studio anywhere','street photographer',
   'rebel without a crew','how not to make a short film','rise of the filmtrepreneur','making movies',
   'the anatomy of story','the art of storytelling','on writing well','writing to learn','how to read book',
   'how to take smart notes','drawing on the right side','constructive anatomy','figure drawing',
   'framed perspective','the story of art','modernism logo','in the blink of an eye','walter murch',
   'best practices','private equity history','quality of earnings','little book of valuation',
   'come into my trading room','japanese candlestick','the most important thing','the new market wizards',
   'poor charlie','the intelligent investor','one up on wall street','learn to earn','stock market genius',
   'economics in one lesson','the ascent of money','anatomy of the bear'],'Reference & Learning'),
 # comics/manga handled separately
 # finance markers
 (['trading','valuation','portfolio','equity research','venture capital','venture deals','angel investing',
   'private equity','hedge fund','wall street','financial shenanigans','accounting','buffett','berkshire',
   'startup','entrepreneur','marketing','advertising','affiliate','copywriting','the goal','pricing',
   'hbr','harvard business'],'Business & Finance'),
 # fitness
 (['bodybuilding','buff dudes','anabolic','starting strength','yoga anatomy','world s fittest','stay fit',
   'poliquin','biohackers','mike mentzer'],'Health & Fitness'),
 # spirituality
 (['bhagavad gita','bhagavata','shiva purana','ramayana','valmiki','mahabharat','gita','purana','vedas','veda',
   'upanishad','bhagats','guru','sikh','buddhist','buddha','tao','zen','yogi','meditation','mantra','tantra',
   'kabir','ashtavakara','panchatantra','jataka'],'Spirituality & Religion'),
 # history
 (['history of','a history','ancient egypt','wounded knee','silk road','the anarchy','world war','the great influenza',
   'rise and fall'],'History'),
 # science
 (['physics','biology','universe','cosmos','quantum','relativity','evolution','the brain','neuroscience','climate',
   'entangled life','astronomy','anatomy','the body','mathematics','number theory'],'Science & Nature'),
 # philosophy
 (['philosophy','nihilism','existential','stoic','consolation','meditations','ethics','metaphysics'],'Philosophy'),
 # poetry
 (['poems','poetry','anthology of','ghazal','shayar','kavya','dohawali','madhushala','gitanjali'],'Poetry & Drama'),
]

# Known individual titles -> genre (normalized), overrides keyword when needed
TITLE_MAP = {
 'atomic habits':'Psychology & Self-Help','the daily stoic':'Psychology & Self-Help','1984':'Classics',
 'animal farm':'Classics','brave new world':'Classics','the handmaid s tale':'Science Fiction',
 'the hunger games':'Young Adult','divergent adult edition':'Young Adult','allegiant':'Young Adult','four':'Young Adult',
 'insurgent adult edition':'Young Adult','the maze runner':'Young Adult','uglies':'Young Adult',
 'life of pi':'Literary Fiction','the alchemist':'Literary Fiction','flowers for algernon':'Science Fiction',
 'the hitchhiker s guide to the galaxy':'Science Fiction','hitchhiker s guide':'Science Fiction',
 'flatland':'Science Fiction','seveneves':'Science Fiction','parable of the sower':'Science Fiction',
 'orbital':'Science Fiction','blindsight':'Science Fiction','the road':'Literary Fiction',
 'never let me go':'Literary Fiction','the memory police':'Literary Fiction','pachinko':'Literary Fiction',
 'circe':'Fantasy','the song of achilles':'Fantasy','galatea':'Fantasy','stardust':'Fantasy','neverwhere':'Fantasy',
 'american gods':'Fantasy','norse mythology':'Fantasy','the graveyard book':'Fantasy','name of the wind':'Fantasy',
 'the name of the wind':'Fantasy','kings of the wyld':'Fantasy','malice':'Fantasy','babel an arcane history':'Fantasy',
 'the name of the rose':'Literary Fiction','baudolino':'Literary Fiction','house of leaves':'Horror',
 'between two fires':'Fantasy','the buffalo hunter hunter':'Horror','uzumaki':'Manga',
 'the forgotten beasts of eld':'Fantasy','starter villain':'Science Fiction','the humans':'Science Fiction',
 'wayward pines series':'Science Fiction','the first fifteen lives of harry august':'Science Fiction',
 'this is how you lose the time war':'Science Fiction','this is how you lose the time war':'Science Fiction',
 'the time traveler s wife':'Romance','the host':'Fantasy','the midnight library':'Literary Fiction',
 'the goldfinch':'Literary Fiction','the secret history':'Literary Fiction','if we were villains':'Mystery & Thriller',
 'gone girl':'Mystery & Thriller','the silent patient':'Mystery & Thriller','silent patient':'Mystery & Thriller',
 'the maltese falcon':'Mystery & Thriller','farewell my lovely':'Mystery & Thriller','labyrinth':'Mystery & Thriller',
 'the queen s gambit':'Literary Fiction','the man who fell to earth':'Science Fiction','mockingbird':'Science Fiction',
 'the hustler':'Literary Fiction','moll flanders':'Classics','the invisible man':'Classics',
 'east of eden':'Classics','the grapes of wrath':'Classics','the pearl':'Classics','the sun also rises':'Classics',
 'tender is the night':'Classics','the great gatsby':'Classics','stoner':'Classics','babbitt':'Classics',
 'atlas shrugged':'Classics','the fountainhead':'Classics','catch 22':'Classics','slaughter house five':'Science Fiction',
 'the catcher in the rye':'Classics','lord of the flies':'Classics','to kill a mockingbird':'Classics',
 'go set a watchman':'Classics','the old man and the sea':'Classics','wuthering heights':'Classics',
 'frankenstein':'Classics','dracula':'Classics','great expectations':'Classics','oliver twist':'Classics',
 'hard times':'Classics','emma':'Classics','pride and prejudice':'Classics','middlemarch':'Classics',
 'mill on the floss':'Classics','the scarlet letter':'Classics','moby dick or the whale':'Classics',
 'heart of darkness':'Classics','the picture of dorarian grey':'Classics','the picture of dorian gray':'Classics',
 'don quixote':'Classics','the three musketeers':'Classics','the count of monte cristo':'Classics',
 'crime and punishment':'Classics','the brothers karamazov':'Classics','the idiot':'Classics','notes from underground':'Classics',
 'the double':'Classics','white nights':'Classics','demons':'Classics','the master and margarita':'Classics',
 'anna karenina':'Classics','war and peace':'Classics','how much land does a man need':'Classics',
 'a calendar of wisdom':'Philosophy','walden':'Philosophy','walden two':'Psychology & Self-Help',
 'the myth of sisyphus and other essays':'Philosophy','thus spoke zarathustra':'Philosophy',
 'beyond good and evil':'Philosophy','on the genealogy of morals and ecce homo':'Philosophy',
 'the consolation of philosophy':'Philosophy','the republic':'Philosophy','republic':'Philosophy',
 'the politics':'Philosophy','meditations':'Philosophy','letters from a stoic':'Philosophy',
 'the story of philosophy':'Philosophy','sophie s world':'Philosophy','the second sex':'Philosophy',
 'either or part i':'Philosophy','at the existentialist cafe':'Philosophy','gravity and grace':'Philosophy',
 'the denial of death':'Philosophy','the problems of philosophy':'Philosophy','probelms of philosophy':'Philosophy',
 'the wisdom of insecurity':'Philosophy','man s search for meaning':'Psychology & Self-Help',
 'thinking fast and slow':'Psychology & Self-Help','the courage to be disliked':'Psychology & Self-Help',
 'a new earth':'Spirituality & Religion','the power of now':'Spirituality & Religion',
 'meditations for self realization':'Spirituality & Religion','the prophet':'Poetry & Drama',
 'paradise lost':'Poetry & Drama','the iliad':'Poetry & Drama','the aeneid':'Poetry & Drama',
 'the epic of gilgamesh':'Poetry & Drama','inferno':'Poetry & Drama','leaves of grass':'Poetry & Drama',
 'the canterbury tales':'Poetry & Drama','canterbury tales':'Poetry & Drama','antigonick':'Poetry & Drama',
 'the odd couple':'Poetry & Drama','slave play':'Poetry & Drama','murder in the cathedral':'Poetry & Drama',
 'the importance of being earnest':'Poetry & Drama','man and superman':'Poetry & Drama',
 'the road to reality':'Science & Nature','six easy pieces':'Science & Nature','the theoretical minimum':'Science & Nature',
 'theoretical minimum guide':'Science & Nature','the elegant universe':'Science & Nature','a brief history of time':'Science & Nature',
 'the universe in a nutshell':'Science & Nature','the theory of everything':'Science & Nature','theory of everything':'Science & Nature',
 'black holes and baby universes and other essays':'Science & Nature','the future of humanity':'Science & Nature',
 'beyond einstein':'Science & Nature','the origin of species':'Science & Nature','origin of species':'Science & Nature',
 'the selfish gene':'Science & Nature','sapiens a brief history of humankind':'History','homo deus':'History',
 'guns germs and steel':'History','the anthropocene reviewed':'Science & Nature','the body a guide for occupants':'Science & Nature',
 'a short history of nearly everything':'Science & Nature','stuff matters':'Science & Nature','your inner fish':'Science & Nature',
 'the language instinct':'Science & Nature','entangled life':'Science & Nature','the road to reality':'Science & Nature',
 'humble pi':'Mathematics','a passion for mathematics':'Mathematics','alex s adventures in numberland':'Mathematics',
 'an introduction to number theory':'Mathematics','the new york times book of mathematics':'Mathematics',
 'mathematical recreations and essays':'Mathematics','the man who knew infinity':'Mathematics','flatland':'Mathematics',
 'the courage to see daily inspiration from great literature':'Reference & Learning',
 'zorba the greek':'Classics','the razor s edge':'Classics','of human bondage':'Classics',
 'the unbearable lightness of being':'Literary Fiction','the wind up bird chronicle':'Literary Fiction',
 'kafka on the shore':'Literary Fiction','norwegian wood':'Literary Fiction','killing commendatore':'Literary Fiction',
 'what i talk about when i talk about running':'Biography & Memoir','the sound and the fury':'Classics',
 'absalom absalom':'Classics','as i lay dying':'Classics','the sound and the fury':'Classics',
 'blood meridian':'Classics','all the king s men':'Classics','the moviegoer':'Classics','company k':'Classics',
 'the painted bird':'Classics','the man without qualities':'Classics','one hundred years of solitude':'Classics',
 'love in the time of cholera':'Classics','no one writes to the colonel':'Classics','by night in chile':'Literary Fiction',
 'labyrinths selected stories and other writings':'Literary Fiction','collected fictions':'Literary Fiction',
 'the collected tales':'Classics','the diary of a madman':'Classics','the life of a stupid man':'Literary Fiction',
 'kappa':'Literary Fiction','in praise of shadows':'Literary Fiction','the narrow road to oku':'Poetry & Drama',
 'sun and steel':'Literary Fiction','confessions of a mask':'Literary Fiction','no longer human':'Literary Fiction',
 'the tale of genji':'Classics','giovanni s room':'Classics','till we have faces':'Fantasy',
 'the chronicles of narnia':'Fantasy','the little white horse':'Children\'s','tales from watership down':'Children\'s',
 'watership down':'Children\'s','flowers for algernon':'Science Fiction','the road':'Literary Fiction',
 'a gentleman in moscow':'Literary Fiction','the book thief':'Literary Fiction','bridge of clay':'Literary Fiction',
 'shantaram':'Literary Fiction','the god of small things':'Literary Fiction','the white tiger':'Literary Fiction',
 'the inheritance of loss':'Literary Fiction','the hungry tide':'Literary Fiction','the living mountain':'Literary Fiction',
 'palace of illusions':'Fantasy','the forest of enchantments':'Fantasy','the last queen':'Literary Fiction',
 'the palace of illusions':'Fantasy','memoirs of a geisha':'Classics','things fall apart':'Classics',
 'the house on mango street':'Literary Fiction','their eyes were watching god':'Classics','beloved':'Literary Fiction',
 'song of solomon':'Literary Fiction','the bluest eye':'Literary Fiction','native son':'Classics',
 'invisible man':'Classics','white teeth':'Literary Fiction','half of a yellow sun':'Literary Fiction',
 'the poppy war':'Fantasy','the starless sea':'Fantasy','piranesi':'Fantasy','the way of kings':'Fantasy',
 'the mistborn trilogy':'Fantasy','the lies of locke lamora':'Fantasy','the gentleman bastard omnibus':'Fantasy',
 'the first law trilogy':'Fantasy','the ember blade':'Fantasy','a darker shade of magic':'Fantasy','vicious':'Fantasy',
 'six of crows':'Fantasy','shadow and bone':'Fantasy','the bear and the nightingale':'Fantasy','the girl in the tower':'Fantasy',
 'his dark materials':'Fantasy','the mortal instruments':'Fantasy','the infernal devices':'Fantasy',
 'lady mightnight':'Fantasy','clockwork princess book 3':'Fantasy','a discovery of witches':'Fantasy',
 'throne of glass':'Fantasy','eragon':'Fantasy','eldest':'Fantasy','the watchmaker of filigree street':'Fantasy',
 'strange weather':'Horror','the ruins':'Horror','the outsider':'Horror','gerald s game':'Horror','insomnia':'Horror',
 'the dark tower 1 the gunslinger':'Horror','books of blood':'Horror','the hellbound heart':'Horror',
 'gutted beautiful horror stories':'Horror','the thief of always':'Horror','the complete works of h p lovecraft':'Horror',
 'scary stories to tell in the dark':'Children\'s','more scary stories to tell in the dark':'Children\'s',
 'the house with a clock in its walls':'Children\'s','bunnicula':'Children\'s','the little prince':'Children\'s',
 'alice s adventures in the wonderland through the looking glass':'Children\'s','the jungle book':'Children\'s',
 'kim':'Classics','the room on the roof':'Literary Fiction','malgudi days':'Literary Fiction',
 'the man eater of malgudi':'Literary Fiction','the financial expert':'Literary Fiction','the english teacher':'Literary Fiction',
 'the village by the sea':'Literary Fiction','baumgartner s bombay':'Literary Fiction','our trees still grow in dehra':'Literary Fiction',
 'a suitable boy':'Literary Fiction','the amazing adventures of kavalier clay':'Literary Fiction',
 'the goldfinch':'Literary Fiction','cloud atlas':'Science Fiction','the poppy war':'Fantasy',
 'a man called ove':'Literary Fiction','the curious incident of the dog in the night time':'Literary Fiction',
 'the perks of being a wallflower':'Young Adult','thirteen reasons why':'Young Adult','the fault in our stars':'Young Adult',
 'looking for alaska':'Young Adult','paper towns':'Young Adult','the hate u give':'Young Adult','fault in our stars':'Young Adult',
 'me before you':'Romance','it ends with us':'Romance','five feet apart':'Romance','when we collided':'Young Adult',
 'the boy in the striped pyjamas':'Young Adult','the book thief':'Literary Fiction','a series of unfortunate events 1 3':'Children\'s',
 'diary of a young girl':'Biography & Memoir','the diary of a young girl':'Biography & Memoir',
 'the anne frank s diary graphic adaptation':'Comics & Graphic Novels','anne frank s diary graphic adaptation':'Comics & Graphic Novels',
 'godfather':'Mystery & Thriller','the godfather':'Mystery & Thriller','the bourne identity':'Mystery & Thriller',
 'the girl with the dragon tattoo':'Mystery & Thriller','angels demons':'Mystery & Thriller','da vinci code':'Mystery & Thriller',
 'the da vinci code':'Mystery & Thriller','sherlock holmes the complete novels and stories vol 1':'Mystery & Thriller',
 'the second opinion':'Mystery & Thriller','and then there were none':'Mystery & Thriller',
 'talking with serial killers':'Mystery & Thriller','and the sea will tell':'Mystery & Thriller',
 'the silmarillion':'Fantasy','the hobbit':'Fantasy','the lord of the rings 1':'Fantasy','the lord of the rings 2':'Fantasy',
 'harry potter the philosophers stone':'Fantasy','fantastic beasts original screenplay 1':'Fantasy',
 'the casual vacancy':'Literary Fiction','the ruin the faithful and the fallen 3':'Fantasy','ruin':'Fantasy',
 'wrath the faithful and the fallen book 4':'Fantasy','the faithful and the fallen 2':'Fantasy',
 's':'Literary Fiction','s by j j abrams and doug dorst':'Literary Fiction','infinite jest':'Literary Fiction',
 'ducks newburyport':'Literary Fiction','no one is talking about this':'Literary Fiction','great circle':'Literary Fiction',
 'we':'Science Fiction','we a novel':'Science Fiction','the man in the high castle':'Science Fiction',
 'time manifold 1':'Science Fiction','the complete robot':'Science Fiction','complete robot':'Science Fiction',
 'i robot film tie in edition':'Science Fiction','prelude to foundation':'Science Fiction','foundation':'Science Fiction',
 'second foundation':'Science Fiction','2010 odyssey two':'Science Fiction','2061 odyssey three':'Science Fiction',
 'dirk gently s holistic detective agency':'Science Fiction','mostly harmless':'Science Fiction',
 'days at the morisaki bookshop':'Literary Fiction','before the coffee gets cold':'Literary Fiction',
 'the memory police':'Literary Fiction','the bridegroom was a dog':'Literary Fiction','the moons of jupiter':'Literary Fiction',
 'the complete stories':'Literary Fiction','can t and won t stories':'Literary Fiction','so late in the day':'Literary Fiction',
 'the collected stories of amy hempel':'Literary Fiction','train dreams':'Literary Fiction','miss lonelyhearts':'Literary Fiction',
 'hurricane season':'Literary Fiction','notes of a crocodile':'Literary Fiction','my year of rest and relaxation':'Literary Fiction',
 'exciting times':'Literary Fiction','normal people':'Literary Fiction','conversations with friends':'Literary Fiction',
 'beautiful world where are you':'Literary Fiction','on earth we re briefly gorgeous':'Literary Fiction',
 'a strangeness in my mind':'Literary Fiction','my name is red':'Literary Fiction','snow':'Literary Fiction',
 'the waves':'Classics','mrs dalloway':'Classics','a room of one s own and three guineas':'Classics',
 'orlando':'Classics','to the lighthouse':'Classics','the hours':'Literary Fiction',
 'the wind up bird chronicle':'Literary Fiction','hamnet':'Literary Fiction','atonement':'Literary Fiction',
 'the humans':'Science Fiction','how to get filthy rich in rising asia':'Literary Fiction',
 'the curious case of benjamin button and other jazz age stories':'Classics','flipped':'Young Adult',
 'aristotle and dante discover the secrets of the universe':'Young Adult','the weight of feathers':'Young Adult',
 'the hating game':'Romance','the best of me':'Romance','the choice':'Romance','the one you cannot have':'Romance',
 'those pricey thakur girls':'Romance','sarah s key':'Literary Fiction','like me':'Literary Fiction',
 'suicide notes':'Young Adult','one of us is lying':'Mystery & Thriller','a discovery of witches':'Fantasy',
 'city of thieves':'Literary Fiction','the painted bird':'Classics','pachinko':'Literary Fiction',
 'the lion women of tehran':'Literary Fiction','the mirror world of melody black':'Literary Fiction',
 'juliet the maniac':'Literary Fiction','a significant life human meaning in a silent universe':'Philosophy',
 'meaning in life and why it matters':'Philosophy','hiking with nietzsche':'Philosophy','the wander society':'Psychology & Self-Help',
 'the book of tea':'Spirituality & Religion','in search of the miraculous':'Spirituality & Religion',
 'krishna the man and his philosophy':'Spirituality & Religion','from sex to superconsciousness':'Spirituality & Religion',
 'the tibetan book of living dying':'Spirituality & Religion','autobiography of a yogi':'Spirituality & Religion',
 'the story of art':'Arts & Film','the letters of vincent van gogh':'Arts & Film','man and his symbols':'Psychology & Self-Help',
 'gone with the wind':'Classics','rebecca':'Classics','jane eyre':'Classics','i claudius':'Classics',
 'the charterhouse of parma':'Classics','the red and the black':'Classics','the sun also rises':'Classics',
 'a farewell to arms':'Classics','for whom the bell tolls':'Classics','across the river and into the trees':'Classics',
 'wide sargasso sea':'Classics','uncle tom s cabin':'Classics','their eyes were watching god':'Classics',
 'the hiding place':'Biography & Memoir','sometimes a great notion':'Classics','the horse s mouth':'Classics',
 'to live a novel':'Literary Fiction','blindness a novel':'Literary Fiction','stoner':'Classics',
 'under the greenwood tree':'Classics','the stone angel':'Classics','the apprenticeship of duddy kravitz':'Classics',
 'north and south':'Classics','brighton rock':'Classics','the heart of the matter':'Classics',
 'the power of full engagement':'Psychology & Self-Help','peak performance':'Psychology & Self-Help',
 'the war of art':'Psychology & Self-Help','so good they can t ignore you':'Psychology & Self-Help',
 'deep work':'Psychology & Self-Help','digital minimalism':'Psychology & Self-Help','slow productivity':'Psychology & Self-Help',
 'four thousand weeks':'Psychology & Self-Help','getting things done':'Psychology & Self-Help','indistractable':'Psychology & Self-Help',
 'the anatomy of story':'Reference & Learning','stumbling on happiness':'Psychology & Self-Help',
 'the happiness hypothesis':'Psychology & Self-Help','the courage to be disliked':'Psychology & Self-Help',
 'algorithms to live by':'Science & Nature','the beginning of infinity':'Science & Nature','the enigma of reason':'Science & Nature',
 'thinking in systems a primer':'Science & Nature','thinking in systems':'Science & Nature','thinking in bets':'Psychology & Self-Help',
 'range':'Psychology & Self-Help','outliers the story of success':'Psychology & Self-Help','talking to strangers':'Psychology & Self-Help',
 'the tipping point':'Psychology & Self-Help','blink':'Psychology & Self-Help','superforecasting':'Psychology & Self-Help',
 'expert political judgment':'Politics & Society','the biggest bluff':'Psychology & Self-Help','educating intuition':'Psychology & Self-Help',
 'great mental models vol 1 general thinking':'Psychology & Self-Help','great mental models vol 2 physics chemistry biology':'Science & Nature',
 'great mental models vol 3 systems mathematics':'Science & Nature','great mental models vol 4 economics art':'Business & Finance',
 'freakonomics':'Business & Finance','superfreakonomics':'Business & Finance','the freakonomics':'Business & Finance',
 'freakonomics hidden side':'Business & Finance','the ascent of money':'Business & Finance','capital in the 21st century':'Business & Finance',
 'the misbehavior of markets':'Business & Finance','lords of finance':'Business & Finance','antifragile':'Business & Finance',
 'fooled by randomness':'Business & Finance','skin in the game':'Business & Finance','the black swan':'Business & Finance',
 'the bed of procrustes':'Philosophy','bed of procrustes':'Philosophy','zero to one':'Business & Finance',
 'shoe dog':'Business & Finance','the everything store':'Business & Finance','working backwards':'Business & Finance',
 'no rules rules':'Business & Finance','good to great':'Business & Finance','built to last':'Business & Finance',
 'the innovator s dilemma':'Business & Finance','crossing the chasm':'Business & Finance','the hard thing about hard things':'Business & Finance',
 'the lean startup':'Business & Finance','the cold start problem':'Business & Finance','blue ocean strategy':'Business & Finance',
 'the 80 20 principle':'Business & Finance','the 48 laws of power':'Psychology & Self-Help','48 laws of power':'Psychology & Self-Help',
 'the art of war':'Politics & Society','the prince':'Politics & Society','the wealth of nations':'Business & Finance',
 'wealth of nations':'Business & Finance','the communist manifesto':'Politics & Society','communist manifesto':'Politics & Society',
 'das kapital':'Politics & Society','manufacturing consent':'Politics & Society','the idea of india':'Politics & Society',
 'the argumentative indian':'Politics & Society','the idea of justice':'Politics & Society','seeing like a state':'Politics & Society',
 'amusing ourselves to death':'Politics & Society','the case against education':'Politics & Society',
 'the anxious generation':'Politics & Society','the madness of crowds':'Politics & Society','so you ve been publicly shamed':'Politics & Society',
 'free speech why it matters':'Politics & Society','on palestine':'Politics & Society','manufacturing consent':'Politics & Society',
 'killing hope':'Politics & Society','inventing reality':'Politics & Society','the dictators handbook':'Politics & Society',
 'social origins of dictatorship and democracy':'Politics & Society','all about love':'Politics & Society',
 'suicide a study in sociology':'Politics & Society','why the poor don t kill us':'Politics & Society',
 'mein kampf':'Politics & Society','the second sex':'Philosophy','why i am an atheist and other essays':'Politics & Society',
 'the god delusion':'Science & Nature','the blind watchmaker':'Science & Nature','a devil s chaplain':'Science & Nature',
 'the ancestor s tale':'Science & Nature','the moral animal':'Science & Nature','behave':'Science & Nature',
 'why we sleep':'Science & Nature','how to change your mind':'Science & Nature','this is your mind on plants':'Science & Nature',
 'the sports gene':'Science & Nature','the tiger':'Science & Nature','the secret wisdom of nature':'Science & Nature',
 'human evolution':'Science & Nature','the great influenza':'Science & Nature','why we run':'Science & Nature',
 'until the end of time':'Science & Nature','five ages universe':'Science & Nature','stuff matters':'Science & Nature',
 'the body keeps the score':'Psychology & Self-Help','maps of meaning':'Psychology & Self-Help',
 '12 rules for life':'Psychology & Self-Help','beyond order':'Psychology & Self-Help','the enigma of reason':'Science & Nature',
 'the elephant in the brain':'Science & Nature','the moral animal':'Science & Nature','stumbling on happiness':'Psychology & Self-Help',
 'quiet':'Psychology & Self-Help','influence the psychology of persuasion':'Psychology & Self-Help',
 'you are not so smart':'Psychology & Self-Help','made to stick':'Business & Finance','never split the difference':'Business & Finance',
 'the tipping point':'Psychology & Self-Help','surrounded by idiots':'Psychology & Self-Help','surrounded by psychopaths':'Psychology & Self-Help',
 'what every body is saying':'Psychology & Self-Help','read people like a book':'Psychology & Self-Help',
 'how to analyze people guide':'Psychology & Self-Help','the game':'Psychology & Self-Help','models':'Psychology & Self-Help',
 'come as you are':'Health & Fitness','she comes first':'Health & Fitness','the man s guide to women':'Health & Fitness',
 'getting the love you want':'Psychology & Self-Help','the course of love':'Philosophy','how to think more about sex':'Philosophy',
 'the consolations of philosophy':'Philosophy','the school of life emotional education':'Psychology & Self-Help',
 'maybe you should talk to someone':'Psychology & Self-Help','how to not die alone':'Psychology & Self-Help',
 'the art of impossible':'Psychology & Self-Help','limitless':'Science Fiction','the productivity project':'Psychology & Self-Help',
 'spark the revolutionary new science of exercise and the brain':'Science & Nature','pivot the only move that matters is your next one':'Psychology & Self-Help',
 'finite and infinite games':'Philosophy','how to take smart notes':'Psychology & Self-Help','feel good productivity':'Psychology & Self-Help',
 'the pathless path':'Psychology & Self-Help','good work reclaiming your inner ambition':'Psychology & Self-Help',
 'essentialism':'Psychology & Self-Help','this essentialism':'Psychology & Self-Help','ego is the enemy':'Psychology & Self-Help',
 'ikigai':'Psychology & Self-Help','the subtle art of not giving a f':'Psychology & Self-Help','can t hurt me':'Psychology & Self-Help',
 'discipline equals freedom field manual':'Psychology & Self-Help','psycho cybernetics':'Psychology & Self-Help',
 'as a man thinketh':'Psychology & Self-Help','the daily stoic':'Psychology & Self-Help','lives of the stoic':'Psychology & Self-Help',
 'transcend the new science of self actualization':'Psychology & Self-Help','the rudest book ever':'Psychology & Self-Help',
 'rudest book ever':'Psychology & Self-Help','101 essays that will change your thinking':'Psychology & Self-Help',
 'personality isn t permanent':'Psychology & Self-Help','how to win friends and influence people in the digital age':'Psychology & Self-Help',
 'rich dad poor dad':'Business & Finance','the psychology of money':'Business & Finance','same as ever':'Business & Finance',
 'poor charlie s almanack':'Business & Finance','poor richard s almanack':'Business & Finance',
 'the intelligent investor':'Business & Finance','one up on wall street':'Business & Finance','learn to earn':'Business & Finance',
 'stock market genius':'Business & Finance','the new market wizards':'Business & Finance','market wizards':'Business & Finance',
 'the essays of warren buffett':'Business & Finance','essays of warren buffett':'Business & Finance',
 'the most important thing illuminated':'Business & Finance','the little book of valuation':'Business & Finance',
 'little book of valuation':'Business & Finance','quality of earnings':'Business & Finance','financial shenanigans':'Business & Finance',
 'financial shenanigans 3rd edition':'Business & Finance','the success equation':'Business & Finance',
 'thinking in bets':'Psychology & Self-Help','the go giver':'Business & Finance','traction':'Business & Finance',
 'the goal a process of ongoing improvement':'Business & Finance','the goal':'Business & Finance',
 'zero to one':'Business & Finance','the cold start problem':'Business & Finance','the 100 startup':'Business & Finance',
 'the 100 dollar startup':'Business & Finance','million dollar weekend':'Business & Finance','never eat alone':'Business & Finance',
 'who the a method for hiring':'Business & Finance','who':'Business & Finance','the pmarca blog':'Business & Finance',
 'venture deals':'Business & Finance','angel investing':'Business & Finance','the innovators':'Business & Finance',
 'chaos monkeys':'Business & Finance','the dumb things smart people do with their money':'Business & Finance',
 'the 80 20 principle':'Business & Finance','tools of titans':'Psychology & Self-Help','the 4 hour work week':'Business & Finance',
 'your money or your life':'Business & Finance','the ascent of money':'Business & Finance',
 'lords of finance':'Business & Finance','the snowball':'Business & Finance','king of capital':'Business & Finance',
 'merger masters':'Business & Finance','dead companies walking':'Business & Finance','one up on wall street':'Business & Finance',
 'competitive strategy':'Business & Finance','confessions of the pricing man':'Business & Finance',
 'the innovator s dilemma':'Business & Finance','the new new thing':'Business & Finance','the partnership':'Business & Finance',
 'for god country and coca cola':'Business & Finance','inside apple':'Business & Finance','alibaba':'Business & Finance',
 'billion dollar loser':'Business & Finance','the everything store':'Business & Finance','only the paranoid survive':'Business & Finance',
 'andrew carnegie':'Biography & Memoir','isaac newton':'Biography & Memoir','einstein his life and universe':'Biography & Memoir',
 'genius the life and science of richard feynman':'Biography & Memoir','surely you re joking mr feynman':'Biography & Memoir',
 'steve jobs the exclusive biography':'Biography & Memoir','elon musk':'Biography & Memoir','iron ambition':'Biography & Memoir',
 'endurance shackleton s incredible voyage':'History','the power broker':'Biography & Memoir','team of rivals':'History',
 'leadership':'History','lyndon johnson the passage of power':'Biography & Memoir','eisenhower in war peace':'Biography & Memoir',
 'boyd':'Biography & Memoir','the fish that ate the whale':'Biography & Memoir','stalin the court of the red tsar':'History',
 'genghis khan and the making of the modern world':'History','a study of history':'History','the story of civilization':'History',
 'the gulag archipelago':'History','one day in the life of ivan denisovich':'Classics','the ghosts of cannae':'History',
 'the archidamian war':'History','the fall of the athenian empire':'History','the rise and fall of ancient egypt':'History',
 'bury my heart at wounded knee':'History','the silk roads':'History','the anarchy':'History','white mughals':'History',
 'city of djinns':'History','the golden road':'History','the last mughal':'History','a history of the sikhs volume 1':'History',
 'sapiens a brief history of humankind':'History','tunnel 29':'History','furious hours':'History',
 'the billionaire raj':'Politics & Society','the polyester prince':'Business & Finance','dongri to dubai':'History',
 'behind the beautiful forevers':'Politics & Society','everybody loves a good drought':'Politics & Society',
 'the great derangement':'Politics & Society','india after gandhi':'History','the discovery of india':'History',
 'the story of my experiments with truth':'Biography & Memoir','wings of fire':'Biography & Memoir',
 'my experiments with truth':'Biography & Memoir','angela s ashes':'Biography & Memoir','this boy s life':'Biography & Memoir',
 'autobiography of a face':'Biography & Memoir','my bondage and my freedom':'Biography & Memoir',
 'the autobiography of malcolm x':'Biography & Memoir','autoboigraphy of malcolm x':'Biography & Memoir',
 'me':'Biography & Memoir','mishima':'Biography & Memoir','stuart':'Biography & Memoir','me elton john':'Biography & Memoir',
 'the tiger':'Science & Nature','the stranger in the woods':'Biography & Memoir','the moth':'Biography & Memoir',
 'educated':'Biography & Memoir','born a crime':'Biography & Memoir','when breath becomes air':'Biography & Memoir',
 'in cold blood':'Mystery & Thriller','the devil in the white city':'History','totto chan':'Biography & Memoir',
 'here there and everywhere':'Biography & Memoir','the daughter from a wishing tree':'Spirituality & Religion',
 'three thousand stitches':'Literary Fiction','house of cards':'Literary Fiction','dear zari':'Biography & Memoir',
 'why i stopped wearing my socks':'Biography & Memoir','making it big':'Biography & Memoir',
 'the billionaire s apprentice':'Business & Finance','losing my virginity':'Business & Finance',
 'direct from dell':'Business & Finance','the snowball':'Business & Finance','rafa my story':'Biography & Memoir',
 'the man who knew infinity':'Mathematics','zen the art of motorcycle maintenance':'Philosophy',
 'facebook inside story':'Business & Finance','daisy jones and the six':'Literary Fiction',
 'unreasonable hospitality':'Business & Finance','the rajneesh chronicles':'Biography & Memoir',
 'the anthropocene reviewed':'Science & Nature','stuff matters':'Science & Nature',
 'i m telling the truth but i m lying essays':'Biography & Memoir','a city on mars':'Science & Nature',
 'hbr':'Business & Finance','harvard business review':'Business & Finance','walden by thoreau':'Philosophy',
 'attention equals life':'Poetry & Drama',
 'the world atlas of coffee':'Reference & Learning','the book of tea':'Spirituality & Religion',
 'in the dust of this planet':'Philosophy','starry speculative corpse':'Philosophy','tentacles longer than night':'Philosophy',
 'flatline constructs':'Philosophy','the divided self':'Philosophy','the politics of experience the bird of paradise':'Philosophy',
 'reason and violence':'Philosophy','nihilism':'Philosophy','nihilism and technology':'Philosophy',
 'utopia':'Philosophy','candide and other stories':'Classics','essays and aphorisms':'Philosophy',
 'the complete essays':'Philosophy','the essential writings of ralph waldo emerson':'Philosophy',
 'the art of happiness':'Philosophy','fragments':'Philosophy','letters from a self made merchant to his son':'Philosophy',
 'the moral sayings of publius syrus':'Philosophy','i think therefore i draw':'Philosophy',
 'the conquest of happiness':'Philosophy','on writing':'Biography & Memoir','on love':'Philosophy',
 'the story of philosophy':'Philosophy','a history of western philosophy':'Philosophy','critique of pure reason':'Philosophy',
 'being and time':'Philosophy','being and nothingness':'Philosophy','the myth of sisyphus':'Philosophy',
 'the stranger':'Classics','the fall':'Classics','the plague':'Classics','nausea':'Classics',
 'tiny beautiful things':'Psychology & Self-Help','bird by bird':'Reference & Learning','on writing well':'Reference & Learning',
 'writing to learn':'Reference & Learning','the elements of style':'Reference & Learning','save the cat':'Reference & Learning',
 'story':'Reference & Learning','the war of art':'Psychology & Self-Help','steal like an artist':'Psychology & Self-Help',
 'show your work':'Psychology & Self-Help','keep going':'Psychology & Self-Help','big magic':'Psychology & Self-Help',
 'daring greatly':'Psychology & Self-Help','the gifts of imperfection':'Psychology & Self-Help',
 'the artist s way':'Psychology & Self-Help','the wander society':'Psychology & Self-Help',
 'numbers don t lie':'Science & Nature','should we eat meat':'Science & Nature','how the world really works':'Science & Nature',
 'factfullness':'Science & Nature','factfulness':'Science & Nature','the rational optimist':'Science & Nature',
 'the moral landscape':'Philosophy','lying':'Philosophy','waking up':'Philosophy','free will':'Philosophy',
 'the sports gene':'Science & Nature','endure':'Health & Fitness','peak':'Psychology & Self-Help',
 'the talent code':'Psychology & Self-Help','mindset':'Psychology & Self-Help','grit':'Psychology & Self-Help',
 'chariots of the gods':'Science & Nature','a brief history of modern india':'History','india unbound':'History',
 'the discovery of india':'History','glimpses of world history':'History',
 # --- second-pass fixes for stragglers ---
 'ramas last act':'Poetry & Drama','ghalib':'Poetry & Drama','the classical tradition of haiku':'Poetry & Drama',
 'the natural wonders':'Science & Nature','everybody lies':'Science & Nature','the bell curve':'Politics & Society',
 'the principles of psychology':'Psychology & Self-Help','the lonesome bodybuilder':'Literary Fiction',
 'sleepwalking':'Literary Fiction','swoon':'Psychology & Self-Help','civil war stories':'Classics',
 'kitab e mirdad':'Spirituality & Religion','lonesome dove':'Classics','story of eye':'Literary Fiction',
 'noopiming':'Literary Fiction','roughing it in the bush':'Classics','against medical advice':'Biography & Memoir',
 'killing and dying stories':'Comics & Graphic Novels','hellfire':'Biography & Memoir','vagabonding':'Reference & Learning',
 'the challenge of pain':'Health & Fitness','bipolar disorder survival guide':'Health & Fitness',
 'the bipolar ii disorder workbook':'Health & Fitness','iit madras':'Reference & Learning','the first 100 course':'Reference & Learning',
 '16 undeniable laws of communication':'Psychology & Self-Help','a year with rilke':'Poetry & Drama',
 'agatha christie miss marple stories':'Mystery & Thriller','all men are mortal':'Classics','only the paranoid survive':'Business & Finance',
 'andrew grove only the paranoid surv':'Business & Finance','blueprint dna us':'Science & Nature','bushido soul of japan':'History',
 'business breakthrough seminar':'Business & Finance','chechnya dagestan history':'History','choose yourself':'Business & Finance',
 'chowringhee':'Literary Fiction','climber mistakes guide':'Health & Fitness','coaching habit change':'Business & Finance',
 'the coaching habit':'Business & Finance','commentaries on living first series':'Spirituality & Religion',
 'dangerous to know':'Fantasy','economy of truth practical maxims':'Business & Finance','experiment without limits':'Psychology & Self-Help',
 'guards guards discworld 8':'Fantasy','hot commodities':'Business & Finance','investment biker':'Business & Finance',
 'how to make the world add up':'Mathematics','how to think a survival guide for a world at odds':'Psychology & Self-Help',
 'india uninc':'Business & Finance','kappa':'Literary Fiction','lifetime cashflow multifamily properties':'Business & Finance',
 'little book of life skills':'Psychology & Self-Help','madness of crowds':'Politics & Society','malice book one':'Fantasy',
 'models revised and updated':'Psychology & Self-Help','nassim taleb skin in':'Business & Finance','skin in the game':'Business & Finance',
 'quirkology big truths':'Psychology & Self-Help','the luck factor':'Psychology & Self-Help','ready fire aim':'Business & Finance',
 'science of being well':'Spirituality & Religion','second sex':'Philosophy','spiritual enlightenment the damnedest thing':'Spirituality & Religion',
 'taxopia the rebel accountant':'Business & Finance','the ceo factory management lessons':'Business & Finance',
 'the ceo factory':'Business & Finance','the long and the short of it':'Business & Finance',
 'the naxalite movement in india':'Politics & Society','the new york times presents smarter by sunday':'Reference & Learning',
 'the new york times science writing':'Science & Nature','the red laugh':'Classics',
 'the unabridged journals of sylvia plath':'Poetry & Drama','the way of the superior man':'Psychology & Self-Help',
 'life at the bottom':'Politics & Society','our culture what s left of it':'Politics & Society',
 'economic facts and fallacies':'Business & Finance','money master':'Business & Finance','tony robbins money master':'Business & Finance',
 'unmasking ai':'Politics & Society','welcome to the jungle':'Business & Finance','what they forgot to teach you at school':'Psychology & Self-Help',
 'world made the west history':'History','hacking box set':'Reference & Learning','the mind of god':'Science & Nature',
 'why has nobody told me this before':'Psychology & Self-Help','marley and me':'Biography & Memoir','miramar':'Classics',
 'what is history':'History','the last of the mohicans':'Classics','the lost river on the trail of the sarasvati':'History',
 'skios a novel':'Literary Fiction','where monsters lie':'Children\'s','speaking tree':'Spirituality & Religion',
 'don t shut up':'Psychology & Self-Help','when they go low we go high':'Politics & Society','corduroy mansions':'Literary Fiction',
 'harry potter and the cursed child':'Fantasy','fantastic beasts original screenplay 2':'Fantasy',
 'encyclopedia of indian history':'History','loonshots':'Business & Finance','if then how the simulmatics corporation invented':'History',
 'quiet the power of introverts':'Psychology & Self-Help','fei fei li the':'Biography & Memoir',
 'the worlds i see':'Biography & Memoir','a significant life':'Philosophy','demons':'Classics',
 'the good the bad and the ugly':'Literary Fiction','wide sargasso sea':'Classics','the second sex':'Philosophy',
}

# extra authors that were missing / caused unclassifieds
AUTHOR_GENRE.update({
 # --- authors added from the Goodreads import ---
 'john perkins':'Politics & Society','liu cixin':'Science Fiction','charlotte perkins':'Classics',
 'h g wells':'Science Fiction','mitch albom':'Biography & Memoir','christopher wilson':'Literary Fiction',
 'ganesh v':'Business & Finance','tahira naqvi':'Literary Fiction','ray bradbury':'Science Fiction',
 'baek se hee':'Psychology & Self-Help','scott h young':'Psychology & Self-Help','zebra learn':'Business & Finance',
 'john boyne':'Young Adult','mohsin hamid':'Literary Fiction','rainbow rowell':'Young Adult',
 'ravinder singh':'Romance','tayari jones':'Literary Fiction','hajime isayama':'Manga','r j palacio':"Children's",
 'ted chiang':'Science Fiction','durjoy datta':'Romance','faith g harper':'Psychology & Self-Help',
 'eric jorgenson':'Business & Finance','tatsuki fujimoto':'Manga','sunil gupta':'Biography & Memoir',
 'chuck palahniuk':'Literary Fiction','courtney summers':'Young Adult','balli kaur jaswal':'Literary Fiction',
 'mary robinette kowal':'Science Fiction','tim collins':"Children's",'seth godin':'Business & Finance',
 'edgar allan poe':'Classics','dav pilkey':"Children's",'arthur miller':'Poetry & Drama','george s clason':'Business & Finance',
 # --- original additions ---
 'j k rowling':'Fantasy','naguib mahfouz':'Classics','e h carr':'History','j fenimore cooper':'Classics',
 'fenimore cooper':'Classics','alexander mccall smith':'Literary Fiction','times of india':'Spirituality & Religion',
 'dr julie smith':'Psychology & Self-Help','julie smith':'Psychology & Self-Help','john grogan':'Biography & Memoir',
 'polly ho yen':'Children\'s','michael frayn':'Literary Fiction','michel donino':'History','michel danino':'History',
 'prakhar gupta':'Psychology & Self-Help','philip collins':'Politics & Society','adrian tomine':'Comics & Graphic Novels',
 'yukiko motoya':'Literary Fiction','meg wolitzer':'Literary Fiction','ambrose bierce':'Classics','larry mcmurtry':'Classics',
 'george bataille':'Literary Fiction','ed macy':'Biography & Memoir','leanne betasamosake':'Literary Fiction',
 'susanna moodie':'Classics','seth stephens':'Science & Nature','rolf potts':'Reference & Learning','bhavabhuti':'Poetry & Drama',
 'prakash pandit':'Poetry & Drama','faubion bowers':'Poetry & Drama','anne lamott':'Reference & Learning',
 'lamott anne':'Reference & Learning','hal friedman':'Biography & Memoir','prioleau betsy':'Psychology & Self-Help',
 'k t davies':'Fantasy','chris sparks':'Psychology & Self-Help','alan jacobs':'Psychology & Self-Help',
 'fei fei li':'Biography & Memoir','rod khleif':'Business & Finance','edwin tenny brewster':'Science & Nature',
 'jay lombard':'Science & Nature','richard j herrnstein':'Politics & Society','charles murray':'Politics & Society',
 'saksham garg':'Literary Fiction','naomi klein':'Politics & Society','simulmatics':'History','jill lepore':'History',
 'dave macleod':'Health & Fitness','k t davies':'Fantasy','anonymous':'Reference & Learning',
})

def find_author_genre(a_norm):
    if not a_norm: return None
    # exact-ish substring match; prefer longer keys
    for key in sorted(AUTHOR_GENRE.keys(), key=len, reverse=True):
        if key in a_norm or a_norm in key:
            # avoid absurd short-key collisions
            if len(key) >= 4:
                return AUTHOR_GENRE[key]
    return None

def find_author_in_title(t_norm):
    """Catch authors embedded in the title text (author field empty)."""
    for key in sorted(AUTHOR_GENRE.keys(), key=len, reverse=True):
        if ' ' in key and len(key) >= 9 and key in t_norm:
            return AUTHOR_GENRE[key]
    return None

def is_manga(t_norm, a_norm):
    if t_norm in MANGA: return True
    if a_norm and a_norm in MANGA_AUTHORS: return True
    return False

def is_comic(t_norm):
    if t_norm in COMIC_TITLES: return True
    for k in COMIC_KEYS:
        if k in t_norm:
            return True
    return False

def classify(b):
    t_norm = norm(b['title'])
    a_norm = norm(b['author'])
    # 1 manga
    if is_manga(t_norm, a_norm):
        return 'Manga'
    # 2 western comic
    if is_comic(t_norm):
        return 'Comics & Graphic Novels'
    # 3 exact title map
    if t_norm in TITLE_MAP:
        return TITLE_MAP[t_norm]
    # 4 author map
    g = find_author_genre(a_norm)
    if g: return g
    # 4b author embedded in the title text
    g = find_author_in_title(t_norm)
    if g: return g
    # 5 title-map partial (substring for series / embedded)  -- before generic keywords
    for key, genre in TITLE_MAP.items():
        if len(key) >= 8 and key in t_norm:
            return genre
    # 6 keyword rules
    for keys, genre in TITLE_KEYWORDS:
        for k in keys:
            if k in t_norm:
                return genre
    return 'Unclassified'

SOURCE_LABEL = {'ebook':'eBook','in my bookshelf':'Bookshelf','goodreads':'Goodreads'}

# --- page-length estimates (drive dot size). If Books TBR.xlsx ever gains a
#     "pages" column, extract.py passes it through and it is used verbatim.
#     Otherwise we estimate from genre + a deterministic per-title variation,
#     with overrides for a few well-known door-stoppers / slim volumes. ---
GENRE_PAGES = {
 'Manga':200,'Comics & Graphic Novels':130,'Young Adult':360,"Children's":130,
 'Fantasy':520,'Science Fiction':360,'Mystery & Thriller':340,'Horror':340,'Romance':340,
 'Literary Fiction':330,'Classics':430,'Poetry & Drama':120,'Science & Nature':360,
 'Mathematics':360,'Philosophy':320,'Psychology & Self-Help':280,'Business & Finance':300,
 'History':460,'Biography & Memoir':400,'Spirituality & Religion':260,'Politics & Society':350,
 'Arts & Film':260,'Health & Fitness':250,'Reference & Learning':420,
}
PAGE_OVERRIDES = {
 'a suitable boy':1349,'war and peace':1225,'infinite jest':1079,'the way of kings':1007,
 'the brothers karamazov':796,'anna karenina':864,'the count of monte cristo':1276,
 'atlas shrugged':1088,'the goldfinch':771,'ducks newburyport':1020,'middlemarch':880,
 'gone with the wind':1037,'the power broker':1246,'the gulag archipelago':660,
 'capital in the 21st century':696,'moby dick or the whale':720,'the lord of the rings 1':531,
 'the name of the wind':662,'shantaram':944,'game of thrones':694,'the mistborn trilogy':720,
 'the wind up bird chronicle':607,'don quixote':1023,'the stand':1153,'ulysses':730,
 'the communist manifesto':48,'the little prince':96,'of mice and men':112,'animal farm':112,
 'the old man and the sea':127,'the great gatsby':180,'the metamorphosis':74,'heart of darkness':96,
 'the prophet':96,'candide and other stories':144,'the stranger':123,'notes from underground':136,
}
def _fnv(s):
    h=2166136261
    for ch in s:
        h ^= ord(ch); h=(h*16777619) & 0xffffffff
    return h
def est_pages(b):
    p=b.get('pages')
    if isinstance(p,(int,float)) and p>0: return int(p)
    key=norm(b['title'])
    if key in PAGE_OVERRIDES: return PAGE_OVERRIDES[key]
    base=GENRE_PAGES.get(b['genre'],320)
    f=0.62 + (_fnv(key)%1000)/1000.0*0.88     # 0.62 .. 1.50
    return int(round(base*f/2)*2)

for i, b in enumerate(books):
    b['id'] = i
    b['genre'] = classify(b)
    b['status'] = 'unread'
    b['format'] = SOURCE_LABEL.get(b.get('source',''), 'eBook')
    b['pages'] = est_pages(b)
    b.pop('source', None)

# ---------------------------------------------------------------------------
# 3. REPORT
# ---------------------------------------------------------------------------
from collections import Counter
gc = Counter(b['genre'] for b in books)
print('=== GENRE DISTRIBUTION ===')
for g,c in gc.most_common():
    print(f'{c:5d}  {g}')
print('total', len(books))
print()
unc = [b for b in books if b['genre']=='Unclassified']
print('=== UNCLASSIFIED (%d) ===' % len(unc))
for b in unc:
    print(f"  {b['title']}  ||  {b['author']}")

slim = [{'id':b['id'],'t':b['title'],'a':b['author'],'g':b['genre'],'f':b['format'],'pg':b.get('pages',300)} for b in books]
data_json = json.dumps(slim, ensure_ascii=False, separators=(',',':')).replace('<','\\u003c')
# Goodreads-derived seed: which books are read + their star ratings (applied once to localStorage)
seed_read = [b['id'] for b in books if b.get('gr_shelf')=='read' or b.get('my_rating')]
seed_ratings = {b['id']: b['my_rating'] for b in books if b.get('my_rating')}
seed_read_json = json.dumps(seed_read, separators=(',',':'))
seed_ratings_json = json.dumps(seed_ratings, separators=(',',':'))

HTML = r'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no">
<title>AV's Library</title>
<style>
:root{
  --bg:#ffffff; --map-bg:#ffffff;
  --ink:#1d2430; --muted:#6b7280; --muted2:#9aa1ad;
  --line:#e7e9ee; --line2:#dfe2e8;
  --panel:#ffffff; --chip:#f4f5f8; --card:#ffffff; --field:#f7f8fb; --hover:#f1f1fb;
  --seg-bg:#f2f3f7; --topbar-bg:rgba(255,255,255,.92); --zoom-bg:rgba(255,255,255,.95); --hint-bg:rgba(255,255,255,.8);
  --accent:#5b57e0; --accent2:#37b6cf; --good:#1aa06d;
  --shadow:0 10px 34px rgba(30,36,60,.14);
  --shadow-sm:0 2px 10px rgba(30,36,60,.10);
  --font:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;
  --topbar-h:56px;
}
body.theme-dark{
  --bg:#05060c; --map-bg:#05060c;
  --ink:#e9edf6; --muted:#9aa2b6; --muted2:#6b7386;
  --line:#1c2236; --line2:#2a3350;
  --panel:#0d111e; --chip:#161c2e; --card:#0d111e; --field:#121728; --hover:#1a2136;
  --seg-bg:#141a2b; --topbar-bg:rgba(9,12,22,.86); --zoom-bg:rgba(16,20,34,.92); --hint-bg:rgba(12,16,28,.7);
  --accent:#8189ff; --accent2:#49c7e0; --good:#38d39b;
  --shadow:0 12px 40px rgba(0,0,0,.55);
  --shadow-sm:0 2px 12px rgba(0,0,0,.4);
}
*{box-sizing:border-box}
html,body{margin:0;height:100%;overflow:hidden;background:var(--bg);color:var(--ink);font-family:var(--font);-webkit-font-smoothing:antialiased}
#app{position:fixed;inset:0}
canvas{display:block;position:absolute;inset:0;top:var(--topbar-h);touch-action:none;cursor:grab;background:var(--map-bg)}
canvas.grabbing{cursor:grabbing}
button{font-family:var(--font)}

/* ---------------- top bar ---------------- */
#topbar{position:absolute;top:0;left:0;right:0;height:var(--topbar-h);z-index:30;display:flex;align-items:center;gap:12px;
  padding:0 14px;background:var(--topbar-bg);backdrop-filter:blur(10px);-webkit-backdrop-filter:blur(10px);
  border-bottom:1px solid var(--line)}
#burger{width:38px;height:38px;border-radius:10px;border:1px solid var(--line2);background:var(--card);cursor:pointer;
  display:flex;flex-direction:column;align-items:center;justify-content:center;gap:4px;flex:none;transition:.15s}
#burger:hover{border-color:var(--accent);background:var(--hover)}
#burger span{width:16px;height:2px;background:var(--ink);border-radius:2px;display:block}
.brand{display:flex;align-items:center;gap:8px;user-select:none}
.brand .logo{width:22px;height:22px;flex:none}
.brand .name{font-weight:800;font-size:17px;letter-spacing:.2px}
.brand .name b{color:var(--accent)}
.brand .tag{font-size:10px;font-weight:700;letter-spacing:1.4px;color:var(--muted2);text-transform:uppercase;margin-top:2px}
#topbar .spacer{flex:1}
#topbar .prog{font-size:12.5px;color:var(--muted);font-variant-numeric:tabular-nums;white-space:nowrap}
#topbar .prog b{color:var(--ink)}
#searchIcon,#themeBtn{width:38px;height:38px;border-radius:10px;border:1px solid var(--line2);background:var(--card);color:var(--ink);cursor:pointer;flex:none;
  display:flex;align-items:center;justify-content:center;transition:.15s}
#searchIcon:hover,#themeBtn:hover{border-color:var(--accent);background:var(--hover);color:var(--accent)}
#themeBtn .ic-sun{display:none}
body.theme-dark #themeBtn .ic-moon{display:none}
body.theme-dark #themeBtn .ic-sun{display:block}

/* ---------------- drawer ---------------- */
#scrim{position:absolute;inset:0;z-index:38;background:rgba(20,24,40,.18);opacity:0;pointer-events:none;transition:.25s}
#scrim.on{opacity:1;pointer-events:auto}
#drawer{position:absolute;top:0;left:0;bottom:0;z-index:40;width:320px;max-width:86vw;background:var(--panel);
  border-right:1px solid var(--line);box-shadow:18px 0 50px rgba(30,36,60,.12);
  transform:translateX(-104%);transition:transform .3s cubic-bezier(.22,.9,.3,1);
  display:flex;flex-direction:column;overflow:hidden}
#drawer.on{transform:translateX(0)}
.dwrap{padding:16px 15px 22px;overflow-y:auto;display:flex;flex-direction:column;gap:13px;height:100%}
.dhead{display:flex;align-items:center;justify-content:space-between}
.dhead .t{font-size:15px;font-weight:800}
.dhead .x{width:30px;height:30px;border:1px solid var(--line2);border-radius:9px;background:var(--card);cursor:pointer;color:var(--muted);font-size:15px}
.dhead .x:hover{color:var(--ink)}

.searchwrap{position:relative}
#search{width:100%;padding:11px 12px 11px 36px;border-radius:11px;border:1px solid var(--line2);background:var(--field);color:var(--ink);
  font-size:14px;font-family:var(--font);outline:none}
#search:focus{border-color:var(--accent);background:var(--card)}
.searchwrap>svg{position:absolute;left:11px;top:50%;transform:translateY(-50%);opacity:.5}
#results{margin-top:7px;max-height:250px;overflow:auto;border-radius:11px;border:1px solid var(--line);background:var(--card);display:none}
#results.on{display:block}
.res{padding:9px 12px;cursor:pointer;border-bottom:1px solid var(--line)}
.res:last-child{border-bottom:none}
.res:hover,.res.sel{background:var(--hover)}
.res .rt{font-size:13px;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.res .ra{font-size:11px;color:var(--muted);margin-top:1px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.res .rg{font-weight:700}

#today{width:100%;padding:12px;border:none;border-radius:12px;cursor:pointer;font-size:14.5px;font-weight:800;color:#fff;letter-spacing:.2px;
  background:linear-gradient(135deg,#6b5cff 0%,#37b6cf 100%);box-shadow:0 6px 18px rgba(91,87,224,.28);
  display:flex;align-items:center;justify-content:center;gap:8px;transition:transform .12s}
#today:hover{transform:translateY(-1px)}

.seg{display:flex;background:var(--seg-bg);border:1px solid var(--line2);border-radius:11px;padding:3px;gap:3px}
.seg button{flex:1;padding:8px 6px;border:none;border-radius:8px;background:transparent;color:var(--muted);font-size:12px;font-weight:700;cursor:pointer;transition:.15s}
.seg button.on{background:var(--card);color:var(--accent);box-shadow:var(--shadow-sm)}

.row{display:flex;gap:8px}
.mini{flex:1;padding:9px 8px;border:1px solid var(--line2);border-radius:10px;background:var(--card);color:var(--muted);font-size:12px;font-weight:700;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:6px;transition:.15s}
.mini:hover{color:var(--ink);border-color:var(--accent)}
.mini.on{color:var(--accent);border-color:var(--accent);background:var(--hover)}

.ldiv{height:1px;background:var(--line);margin:3px 0}
.lh{display:flex;align-items:center;justify-content:space-between;padding:0 2px}
.lh span{font-size:11px;font-weight:800;color:var(--muted);text-transform:uppercase;letter-spacing:.7px}
.lh a{font-size:11px;color:var(--accent);cursor:pointer}
#genres{display:flex;flex-direction:column;gap:1px}
.gitem{display:flex;align-items:center;gap:9px;padding:6px 7px;border-radius:9px;cursor:pointer;transition:.12s}
.gitem:hover{background:var(--hover)}
.gitem.off{opacity:.4}
.dot{width:10px;height:10px;border-radius:50%;flex:none}
.gitem .gname{flex:1;font-size:12.5px;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.gitem .gcount{font-size:11px;color:var(--muted2);font-variant-numeric:tabular-nums}

/* ---------------- zoom ---------------- */
#zoom{position:absolute;bottom:20px;right:16px;z-index:22;display:flex;flex-direction:column;gap:8px}
#zoom button{width:40px;height:40px;border-radius:11px;border:1px solid var(--line2);background:var(--zoom-bg);color:var(--ink);
  font-size:19px;cursor:pointer;box-shadow:var(--shadow-sm);display:flex;align-items:center;justify-content:center;transition:.12s}
#zoom button:hover{border-color:var(--accent);color:var(--accent)}
#zoom .z-lbl{font-size:11px;font-weight:700}

#hint{position:absolute;bottom:22px;left:50%;transform:translateX(-50%);z-index:15;font-size:12px;color:var(--muted2);
  background:var(--hint-bg);padding:6px 14px;border-radius:20px;border:1px solid var(--line);pointer-events:none;transition:opacity .6s;white-space:nowrap}

/* ---------------- tooltip (dark pill) ---------------- */
#tip{position:absolute;z-index:44;pointer-events:none;display:none;max-width:250px;padding:8px 11px;background:#1e2330;color:#fff;
  border-radius:10px;box-shadow:0 8px 24px rgba(20,24,40,.28);transform:translate(-50%,calc(-100% - 14px))}
#tip .tt{font-size:12.5px;font-weight:700;line-height:1.25}
#tip .ta{font-size:11px;color:#c3c7d4;margin-top:2px}
#tip .tg{font-size:10px;margin-top:5px;display:inline-flex;align-items:center;gap:5px;font-weight:700}
#tip .ts{font-size:10px;color:#9aa1b2;margin-top:3px}

/* ---------------- detail ---------------- */
#detail{position:absolute;top:var(--topbar-h);right:0;bottom:0;z-index:36;width:350px;max-width:88vw;background:var(--card);
  border-left:1px solid var(--line);box-shadow:-16px 0 46px rgba(30,36,60,.14);
  transform:translateX(105%);transition:transform .32s cubic-bezier(.22,.9,.3,1);display:flex;flex-direction:column;padding:22px;overflow-y:auto}
#detail.on{transform:translateX(0)}
#detail .close{position:absolute;top:15px;right:15px;width:31px;height:31px;border-radius:9px;border:1px solid var(--line2);background:var(--card);color:var(--muted);cursor:pointer;font-size:15px}
#detail .close:hover{color:var(--ink)}
#detail .dtag{align-self:flex-start;margin-top:6px;font-size:11px;font-weight:800;text-transform:uppercase;letter-spacing:.6px;padding:5px 11px;border-radius:20px;display:inline-flex;align-items:center;gap:6px}
#detail h2{margin:15px 0 5px;font-size:22px;line-height:1.2;font-weight:800}
#detail .dauth{font-size:14px;color:var(--muted);margin-bottom:2px}
#detail .dmeta{margin-top:16px;display:flex;flex-direction:column;gap:9px}
#detail .drow{display:flex;justify-content:space-between;font-size:13px;padding:9px 12px;background:var(--field);border-radius:10px;border:1px solid var(--line)}
#detail .drow b{color:var(--muted);font-weight:600}
#markbtn{margin-top:18px;width:100%;padding:14px;border:none;border-radius:12px;cursor:pointer;font-size:15px;font-weight:800;color:#fff;background:linear-gradient(135deg,#1aa06d,#159f8e);transition:.15s;display:flex;align-items:center;justify-content:center;gap:8px}
#markbtn:hover{filter:brightness(1.05)}
#markbtn.isread{background:var(--seg-bg);color:var(--muted);border:1px solid var(--line2)}
#detail .dsub{margin-top:20px;font-size:11px;font-weight:800;text-transform:uppercase;letter-spacing:.6px;color:var(--muted2)}
#detail .chips{display:flex;flex-wrap:wrap;gap:7px;margin-top:10px}
#detail .chip{font-size:12px;padding:7px 11px;border-radius:9px;background:var(--chip);border:1px solid var(--line);cursor:pointer;color:var(--muted);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:100%;transition:.12s}
#detail .chip:hover{color:var(--ink);border-color:var(--accent)}

/* ---------------- star ratings ---------------- */
.stars{display:inline-flex;gap:2px;align-items:center}
.stars .star{cursor:pointer;font-size:22px;line-height:1;color:var(--line2);transition:color .1s,transform .1s;user-select:none}
.stars .star:hover{transform:scale(1.14)}
.stars .star.on{color:#f5b301}
.stars.ro{gap:1px}
.stars.ro .star{cursor:default;font-size:13px}
.stars.ro .star:hover{transform:none}
#detail .drate{margin-top:16px;display:flex;align-items:center;gap:10px;padding:11px 12px;background:var(--field);border:1px solid var(--line);border-radius:11px}
#detail .drate .dlabel{font-size:11px;font-weight:800;color:var(--muted);text-transform:uppercase;letter-spacing:.6px}
#detail .drate .dratev{font-size:12px;color:var(--muted2);margin-left:auto;white-space:nowrap}
.ratefilter{display:flex;align-items:center;gap:8px;background:var(--seg-bg);border:1px solid var(--line2);border-radius:11px;padding:6px 10px}
.ratefilter .rf-label{font-size:12px;font-weight:700;color:var(--muted);white-space:nowrap}
.ratefilter .stars{flex:1}
.ratefilter .stars .star{font-size:19px}
.ratefilter .rf-clear{border:none;background:transparent;color:var(--muted2);cursor:pointer;font-size:15px;padding:2px 5px;border-radius:6px;visibility:hidden}
.ratefilter .rf-clear.on{visibility:visible}
.ratefilter .rf-clear:hover{color:var(--ink)}
#topbar .prog .ravg{color:#e0a300;font-weight:700}
.res .rr{color:#f5b301}
#tip .trate{color:#ffcf4d;font-size:13px;margin-top:4px;letter-spacing:1px}

/* ---------------- modal ---------------- */
.overlay{position:absolute;inset:0;z-index:60;background:rgba(20,24,40,.28);display:none;align-items:center;justify-content:center;padding:20px}
.overlay.on{display:flex}
.modal{width:520px;max-width:100%;max-height:86vh;overflow:auto;padding:24px;border-radius:16px;background:var(--card);box-shadow:var(--shadow)}
.modal h3{margin:0 0 6px;font-size:19px}
.modal p{margin:0 0 14px;font-size:13px;color:var(--muted);line-height:1.5}
.modal textarea{width:100%;height:190px;resize:vertical;border-radius:11px;border:1px solid var(--line2);background:var(--field);color:var(--ink);font-family:var(--font);font-size:13px;padding:12px;outline:none}
.modal textarea:focus{border-color:var(--accent);background:var(--card)}
.modal .mrow{display:flex;gap:10px;margin-top:14px}
.modal .btn{flex:1;padding:12px;border:none;border-radius:11px;font-size:14px;font-weight:800;cursor:pointer}
.btn-primary{background:linear-gradient(135deg,#6b5cff,#37b6cf);color:#fff}
.btn-ghost{background:var(--seg-bg);color:var(--muted);border:1px solid var(--line2)!important}
#importResult{margin-top:12px;font-size:13px;color:var(--good);display:none}

#toast{position:absolute;bottom:74px;left:50%;transform:translateX(-50%) translateY(16px);z-index:70;background:#1e2330;color:#fff;
  border-radius:11px;padding:11px 18px;font-size:13.5px;font-weight:600;box-shadow:0 10px 30px rgba(20,24,40,.3);opacity:0;transition:.3s;pointer-events:none;max-width:80vw;text-align:center}
#toast.on{opacity:1;transform:translateX(-50%) translateY(0)}

.dwrap::-webkit-scrollbar,#detail::-webkit-scrollbar,#results::-webkit-scrollbar{width:8px}
.dwrap::-webkit-scrollbar-thumb,#detail::-webkit-scrollbar-thumb,#results::-webkit-scrollbar-thumb{background:#d7dae1;border-radius:8px}
@media (max-width:560px){ #topbar .prog{display:none} }
</style>
</head>
<body>
<div id="app">
  <canvas id="cv"></canvas>

  <header id="topbar">
    <button id="burger" aria-label="Menu"><span></span><span></span><span></span></button>
    <div class="brand">
      <svg class="logo" viewBox="0 0 24 24"><path d="M12 2l2.4 6.6L21 11l-6.6 2.4L12 20l-2.4-6.6L3 11l6.6-2.4L12 2z" fill="url(#lg)"/><defs><linearGradient id="lg" x1="3" y1="2" x2="21" y2="20"><stop stop-color="#6b5cff"/><stop offset="1" stop-color="#37b6cf"/></linearGradient></defs></svg>
      <div><div class="name">AV's <b>Library</b></div></div>
    </div>
    <div class="spacer"></div>
    <div class="prog"><b id="pcount">0</b> / <span id="tcount">0</span> read<span id="ravgWrap"></span></div>
    <button id="themeBtn" aria-label="Toggle light or dark" title="Toggle light / dark">
      <svg class="ic-moon" width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12.8A9 9 0 1 1 11.2 3 7 7 0 0 0 21 12.8z"/></svg>
      <svg class="ic-sun" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/></svg>
    </button>
    <button id="searchIcon" aria-label="Search"><svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="7"/><path d="M21 21l-4.3-4.3"/></svg></button>
  </header>

  <div id="scrim"></div>
  <aside id="drawer">
    <div class="dwrap">
      <div class="dhead"><div class="t">Explore</div><button class="x" id="drawerClose">✕</button></div>
      <div class="searchwrap">
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#9aa1ad" stroke-width="2"><circle cx="11" cy="11" r="7"/><path d="M21 21l-4.3-4.3"/></svg>
        <input id="search" placeholder="Search title or author…" autocomplete="off" spellcheck="false">
        <div id="results"></div>
      </div>
      <button id="today">✦ Read Today</button>
      <div class="seg" id="statusSeg">
        <button data-s="all" class="on">All</button>
        <button data-s="unread">Unread</button>
        <button data-s="read">Read</button>
      </div>
      <div class="ratefilter">
        <span class="rf-label">Rating ≥</span>
        <span class="stars" id="ratefilterStars"></span>
        <button class="rf-clear" id="rateClear" title="Clear rating filter">✕</button>
      </div>
      <div class="row">
        <button class="mini on" id="fmtAll" data-fmt="all">All formats</button>
        <button class="mini" id="fmtE" data-fmt="eBook">eBook</button>
        <button class="mini" id="fmtB" data-fmt="Bookshelf">Shelf</button>
      </div>
      <div class="row">
        <button class="mini" id="connBtn">✧ Connections</button>
      </div>
      <div class="ldiv"></div>
      <div class="lh"><span>Genres</span><a id="genAll">reset</a></div>
      <div id="genres"></div>
    </div>
  </aside>

  <div id="zoom">
    <button id="zin">+</button>
    <button id="zout">−</button>
    <button id="zfit" title="Fit to view"><span class="z-lbl">fit</span></button>
  </div>

  <div id="hint">Drag to pan · scroll to zoom · hover a region for its name · click a dot for details · ☰ menu</div>
  <div id="tip"></div>

  <div id="detail">
    <button class="close" id="dclose">✕</button>
    <div class="dtag" id="dtag"></div>
    <h2 id="dtitle"></h2>
    <div class="dauth" id="dauth"></div>
    <div class="dmeta">
      <div class="drow"><b>Genre</b><span id="dgenre"></span></div>
      <div class="drow"><b>Format</b><span id="dformat"></span></div>
      <div class="drow"><b>Status</b><span id="dstatus"></span></div>
    </div>
    <div class="drate">
      <span class="dlabel">Your rating</span>
      <span class="stars" id="dstars"></span>
      <span class="dratev" id="dratev">Not rated</span>
    </div>
    <button id="markbtn"></button>
    <div class="dsub" id="dbysub" style="display:none">More by this author</div>
    <div class="chips" id="dbyauthor"></div>
    <div class="dsub">Explore this genre</div>
    <div class="chips" id="dgenrechips"></div>
  </div>

  <div id="toast"></div>
</div>

<script>const BOOKS = __BOOKS_JSON__;
const SEED_READ = __SEED_READ__;        /* Goodreads: ids marked read */
const SEED_RATINGS = __SEED_RATINGS__;  /* Goodreads: {id: 1-5 stars} */</script>
<script>
(function(){
"use strict";
var LS_KEY="avs-library-read-v1";
var GOLDEN=Math.PI*(3-Math.sqrt(5));
// DARK-MODE "cosmic" palette (vivid stars on black); also used for the legend swatches
var GENRE_COLORS={
 "Manga":"#ffffff", "Comics & Graphic Novels":"#f0f8ff", "Young Adult":"#e6f2ff", "Children's":"#cce6ff",
 "Fantasy":"#b3d9ff", "Science Fiction":"#00ffff", "Mystery & Thriller":"#80bfff", "Horror":"#66b3ff",
 "Romance":"#4da6ff", "Literary Fiction":"#3399ff", "Classics":"#1a8cff", "Poetry & Drama":"#0080ff",
 "Science & Nature":"#0073e6", "Mathematics":"#0066cc", "Philosophy":"#0059b3", "Psychology & Self-Help":"#004d99",
 "Business & Finance":"#004080", "History":"#003366", "Biography & Memoir":"#00264d", "Spirituality & Religion":"#e0ffff",
 "Politics & Society":"#b3ffff", "Arts & Film":"#80ffff", "Health & Fitness":"#4dffff", "Reference & Learning":"#00e6e6"
};
function colorFor(g){ return GENRE_COLORS[g]||"#a8b6d6"; }

// LIGHT MODE: neurons use genre color, but overly white colors become grey so the cloud/glow isn't too heavy
function getLightModeColor(hex){
  if(!hex) return "#777788";
  var h = hex.replace("#","");
  var r = parseInt(h.substr(0,2),16);
  var g = parseInt(h.substr(2,2),16);
  var b = parseInt(h.substr(4,2),16);
  if(r+g+b > 720) return "#777788"; 
  return hex;
}
function dotColor(b){
  if(theme==="light") return getLightModeColor(b.color);
  return b.color;
}
// page count -> dot radius (world units); sqrt so area tracks length, clamped
function pagesToRad(pg){ pg=pg||300; return Math.max(1.1, Math.min(6.6, 0.55+0.135*Math.sqrt(pg))); }

var readSet=loadRead();
var RATE_KEY="avs-library-ratings-v1";
var ratings=(function(){ try{ return JSON.parse(localStorage.getItem(RATE_KEY)||"{}")||{}; }catch(e){ return {}; } })();
var THEME_KEY="avs-library-theme-v1";
var theme=(function(){ try{ return localStorage.getItem(THEME_KEY)||"light"; }catch(e){ return "light"; } })();
var STRETCHX=1.32;
var state={status:"all",format:"all",search:"",hidden:{},showConn:false,selected:null,hover:null,hoverGenre:null,minRating:0};
var cam={x:0,y:0,s:1,tx:0,ty:0,ts:1,fx:0,fy:0,fs:1,animating:false,animStart:0,animDur:0};

function loadRead(){ try{var r=JSON.parse(localStorage.getItem(LS_KEY)||"[]");var s={};r.forEach(function(id){s[id]=1;});return s;}catch(e){return {};} }
function saveRead(){ try{localStorage.setItem(LS_KEY,JSON.stringify(Object.keys(readSet).map(Number)));}catch(e){} }
function isRead(id){ return !!readSet[id]; }
// ---- ratings (Goodreads-style 1-5 stars; rating a book implies it's read) ----
function saveRatings(){ try{localStorage.setItem(RATE_KEY,JSON.stringify(ratings));}catch(e){} }
function getRating(id){ return ratings[id]||0; }
function setRating(id,r){
  if(r>0){ ratings[id]=r; readSet[id]=1; }        // rating marks the book read
  else { delete ratings[id]; }                     // clearing rating keeps read status
  saveRatings(); saveRead();
}
function starStr(n){ var s=""; for(var i=1;i<=5;i++) s+=(i<=n?"★":""); return s; }  // filled only, for compact display
function mul32(a){ return function(){ a|=0;a=a+0x6D2B79F5|0;var t=Math.imul(a^a>>>15,1|a);t=t+Math.imul(t^t>>>7,61|t)^t;return ((t^t>>>14)>>>0)/4294967296; }; }
function hashStr(s){ var h=2166136261;for(var i=0;i<s.length;i++){h^=s.charCodeAt(i);h=Math.imul(h,16777619);}return h>>>0; }
function norm(s){ return (s||"").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g,"").replace(/[^a-z0-9 ]/g," ").replace(/\s+/g," ").trim(); }
function gauss(rng){ var u=Math.max(1e-6,rng()),v=rng(); return Math.sqrt(-2*Math.log(u))*Math.cos(2*Math.PI*v); }

// ---------------- layout ----------------
var genres=[], genreMeta={}, edges=[], byId={};
BOOKS.forEach(function(b){ byId[b.id]=b; });
function buildLayout(){
  var counts={}; BOOKS.forEach(function(b){ counts[b.g]=(counts[b.g]||0)+1; });
  genres=Object.keys(counts).sort(function(a,b){ return counts[b]-counts[a]; });
  var placed=[], SP=18, OVER=0.70;   // OVER<1 lets clusters overlap -> connected landmass (lower = tighter)
  genres.forEach(function(g){
    var n=counts[g], r=31*Math.sqrt(n)+10, k=0, pt;
    while(true){
      var rad=SP*Math.sqrt(k+1), ang=k*GOLDEN; pt={x:rad*Math.cos(ang),y:rad*Math.sin(ang)};
      var ok=true;
      for(var i=0;i<placed.length;i++){ var p=placed[i],dx=pt.x-p.x,dy=pt.y-p.y; if(Math.sqrt(dx*dx+dy*dy)<(r+p.r)*OVER){ok=false;break;} }
      if(ok) break; k++; if(k>300000) break;
    }
    placed.push({x:pt.x,y:pt.y,r:r});
    genreMeta[g]={cx:pt.x,cy:pt.y,r:r,n:n,color:colorFor(g)};
  });
  var byGenre={}; BOOKS.forEach(function(b){ (byGenre[b.g]=byGenre[b.g]||[]).push(b); });
  genres.forEach(function(g){
    var arr=byGenre[g], m=genreMeta[g], R=m.r, rng=mul32(hashStr(g));
    arr.sort(function(a,b){ var aa=(a.a||"~").toLowerCase(),ba=(b.a||"~").toLowerCase(); if(aa!==ba)return aa<ba?-1:1; return a.t.toLowerCase()<b.t.toLowerCase()?-1:1; });
    // group consecutive same-author (non-empty) into clumps; empty-author books are singletons
    var i=0;
    while(i<arr.length){
      var a=norm(arr[i].a), grp=[arr[i]], j=i+1;
      if(a.length>=4){ while(j<arr.length && norm(arr[j].a)===a){ grp.push(arr[j]); j++; } }
      // seed position within disc, denser toward centre
      var u=rng(), rr=R*Math.pow(u,0.55)*0.92, sang=rng()*Math.PI*2;
      var sx=m.cx+rr*Math.cos(sang), sy=m.cy+rr*Math.sin(sang);
      var spread=Math.min(R*0.17, R*0.04+Math.sqrt(grp.length)*R*0.022);
      for(var q=0;q<grp.length;q++){
        var b=grp[q];
        var x=sx+gauss(rng)*spread, y=sy+gauss(rng)*spread;
        var dx=x-m.cx, dy=y-m.cy, d=Math.sqrt(dx*dx+dy*dy), lim=R*1.05;
        if(d>lim){ x=m.cx+dx/d*lim; y=m.cy+dy/d*lim; }
        b.x=x; b.y=y; b.color=m.color;
        b.rad=pagesToRad(b.pg);   // dot size reflects the book's page length
      }
      i=j;
    }
  });
  // horizontal stretch so the field fills landscape screens (like the reference map)
  BOOKS.forEach(function(b){ b.x*=STRETCHX; });
  genres.forEach(function(g){ genreMeta[g].cx*=STRETCHX; });
  // constellation edges based on genres (neural network look)
  var byGenre={};
  BOOKS.forEach(function(b){ (byGenre[b.g]=byGenre[b.g]||[]).push(b); });
  Object.keys(byGenre).forEach(function(g){ 
    var arr=byGenre[g]; 
    if(arr.length<2) return; 
    arr.sort(function(a,b){ return a.x - b.x; });
    for(var i=0; i<arr.length-1; i++){
      edges.push([arr[i], arr[i+1]]);
      if(i+2 < arr.length && Math.random() < 0.6) edges.push([arr[i], arr[i+2]]);
      if(i+3 < arr.length && Math.random() < 0.3) edges.push([arr[i], arr[i+3]]);
    }
  });
}

// ---------------- canvas ----------------
var cv=document.getElementById("cv"), ctx=cv.getContext("2d");
var DPR=Math.min(window.devicePixelRatio||1,2), W=0,H=0, TOP=56;
function resize(){
  W=window.innerWidth; H=window.innerHeight-TOP;
  cv.width=W*DPR; cv.height=H*DPR; cv.style.width=W+"px"; cv.style.height=H+"px";
  ctx.setTransform(DPR,0,0,DPR,0,0);
  if(!userMoved) fitView(0);
  requestDraw();
}
window.addEventListener("resize",resize);

function w2s(x,y){ return [ (x-cam.x)*cam.s + W/2, (y-cam.y)*cam.s + H/2 ]; }
function s2w(px,py){ return [ (px-W/2)/cam.s + cam.x, (py-H/2)/cam.s + cam.y ]; }
function hexA(hex,a){ var h=hex.replace("#","");return "rgba("+parseInt(h.substr(0,2),16)+","+parseInt(h.substr(2,2),16)+","+parseInt(h.substr(4,2),16)+","+a+")"; }

// DARK-mode additive glow sprite per colour (stars shine on black)
var glowCache={};
function glowSprite(color){
  if(glowCache[color]) return glowCache[color];
  var s=128, c=document.createElement("canvas"); c.width=s; c.height=s; var g=c.getContext("2d");
  var grd=g.createRadialGradient(s/2,s/2,0,s/2,s/2,s/2);
  grd.addColorStop(0,color);
  grd.addColorStop(0.15,color); 
  grd.addColorStop(0.4,hexA(color,0.6));
  grd.addColorStop(1,hexA(color,0));
  g.fillStyle=grd; g.fillRect(0,0,s,s); glowCache[color]=c; return c;
}
// LIGHT-mode soft grey "plate bloom" for large dots (astrophoto look)
var bloomSprite=(function(){
  var s=64, c=document.createElement("canvas"); c.width=s; c.height=s; var g=c.getContext("2d");
  var grd=g.createRadialGradient(s/2,s/2,0,s/2,s/2,s/2);
  grd.addColorStop(0,"rgba(18,26,44,.5)"); grd.addColorStop(0.5,"rgba(18,26,44,.12)"); grd.addColorStop(1,"rgba(18,26,44,0)");
  g.fillStyle=grd; g.fillRect(0,0,s,s); return c;
})();
// faint static background starfield for dark mode (screen-space, deterministic)
var BGSTARS=(function(){ var a=[], rng=mul32(20240808); for(var i=0;i<220;i++){ a.push([rng(),rng(),0.4+rng()*1.2,0.15+rng()*0.5]); } return a; })();
function drawBgStars(){ ctx.fillStyle="#cfe0ff"; for(var i=0;i<BGSTARS.length;i++){ var s=BGSTARS[i]; ctx.globalAlpha=s[3]*0.5; ctx.beginPath(); ctx.arc(s[0]*W,s[1]*H,s[2],0,7); ctx.fill(); } ctx.globalAlpha=1; }

function passFilter(b){
  if(state.hidden[b.g]) return false;
  if(state.format!=="all" && b.f!==state.format) return false;
  if(state.status==="read" && !isRead(b.id)) return false;
  if(state.status==="unread" && isRead(b.id)) return false;
  if(state.minRating>0 && getRating(b.id)<state.minRating) return false;   // rating filter
  if(state.search && (b.t+" "+b.a).toLowerCase().indexOf(state.search)<0) return false;
  return true;
}

var drawPending=false;
function requestDraw(){ if(!drawPending){ drawPending=true; requestAnimationFrame(draw); } }
function draw(ts){
  drawPending=false;
  if(!isFinite(cam.s)||cam.s<=0) cam.s=0.1; if(!isFinite(cam.x)) cam.x=0; if(!isFinite(cam.y)) cam.y=0;
  if(W<=0||H<=0) return;
  if(cam.animating){
    var t=Math.min(1,(ts-cam.animStart)/cam.animDur);
    var e=t<.5?4*t*t*t:1-Math.pow(-2*t+2,3)/2;
    cam.x=cam.fx+(cam.tx-cam.fx)*e; cam.y=cam.fy+(cam.ty-cam.fy)*e; cam.s=cam.fs+(cam.ts-cam.fs)*e;
    if(t>=1) cam.animating=false; requestDraw();
  }
  var dark = theme==="dark";
  ctx.clearRect(0,0,W,H);
  if(dark){
    var bg=ctx.createRadialGradient(W*0.5,H*0.42,0,W*0.5,H*0.42,Math.max(W,H)*0.82);
    bg.addColorStop(0,"#0b1124"); bg.addColorStop(0.55,"#070912"); bg.addColorStop(1,"#04050b");
    ctx.fillStyle=bg; ctx.fillRect(0,0,W,H);
    drawBgStars();
  } else {
    ctx.fillStyle="#ffffff"; ctx.fillRect(0,0,W,H);
  }

  // galaxy background clouds per genre
  ctx.globalCompositeOperation=dark?"lighter":"source-over";
  genres.forEach(function(g){
    if(state.hidden[g]) return;
    var m=genreMeta[g], sp=w2s(m.cx,m.cy), cr=m.r*cam.s*1.8;
    if(sp[0]<-cr||sp[0]>W+cr||sp[1]<-cr||sp[1]>H+cr) return;
    var grd=ctx.createRadialGradient(sp[0],sp[1],0,sp[0],sp[1],cr);
    var col=m.color;
    if(!dark) col = getLightModeColor(col);
    var isHover = (state.hoverGenre === g);
    if(dark){
      grd.addColorStop(0,hexA(col,isHover?0.40:0.20));
      grd.addColorStop(0.4,hexA(col,isHover?0.12:0.06));
      grd.addColorStop(1,hexA(col,0));
    } else {
      grd.addColorStop(0,hexA(col,isHover?0.35:0.22));
      grd.addColorStop(0.4,hexA(col,isHover?0.20:0.14));
      grd.addColorStop(1,hexA(col,0));
    }
    ctx.fillStyle=grd; ctx.beginPath(); ctx.arc(sp[0],sp[1],cr,0,7); ctx.fill();
  });
  ctx.globalCompositeOperation="source-over";

  // luminous pass: additive star glow (dark) OR neuron glow (light)
  if(dark){
    ctx.globalCompositeOperation="lighter";
    for(var k=0;k<BOOKS.length;k++){
      var b=BOOKS[k]; if(!passFilter(b)) continue;
      var sp2=w2s(b.x,b.y); var rr=Math.max(1.2,b.rad*cam.s); var gz=Math.max(30, rr*22);
      if(sp2[0]<-gz||sp2[0]>W+gz||sp2[1]<-gz||sp2[1]>H+gz) continue;
      ctx.globalAlpha=isRead(b.id)&&state.status!=="read"?0.04:0.12;
      ctx.drawImage(glowSprite(b.color), sp2[0]-gz/2, sp2[1]-gz/2, gz, gz);
    }
    ctx.globalAlpha=1; ctx.globalCompositeOperation="source-over";
  } else {
    for(var k=0;k<BOOKS.length;k++){
      var b=BOOKS[k]; if(!passFilter(b)) continue;
      var rr=Math.max(1.2,b.rad*cam.s);
      var sp2=w2s(b.x,b.y); var bz=Math.max(30, rr*22);
      if(sp2[0]<-bz||sp2[0]>W+bz||sp2[1]<-bz||sp2[1]>H+bz) continue;
      ctx.globalAlpha=isRead(b.id)&&state.status!=="read"?0.12:0.45; 
      ctx.drawImage(glowSprite(dotColor(b)), sp2[0]-bz/2, sp2[1]-bz/2, bz, bz);
    }
    ctx.globalAlpha=1;
  }

  // optional faint connections
  if((!dark || state.showConn) && cam.s>0.05){
    ctx.lineWidth=dark ? Math.min(2.5,0.8*cam.s+0.4) : Math.min(1.8,0.6*cam.s+0.2);
    for(var i=0;i<edges.length;i++){
      var a=edges[i][0], b=edges[i][1], pa=passFilter(a), pb=passFilter(b);
      if(!pa&&!pb) continue;
      var s1=w2s(a.x,a.y), s2=w2s(b.x,b.y);
      if((s1[0]<0&&s2[0]<0)||(s1[0]>W&&s2[0]>W)||(s1[1]<0&&s2[1]<0)||(s1[1]>H&&s2[1]>H)) continue;
      var sel=state.selected&&(a.id===state.selected||b.id===state.selected);
      var alpha = dark ? ((pa&&pb)?0.25:0.05) : ((pa&&pb)?0.30:0.08); 
      ctx.strokeStyle=sel?hexA(a.color,.8):(dark ? hexA(a.color,alpha) : hexA(dotColor(a), alpha));
      ctx.beginPath(); ctx.moveTo(s1[0],s1[1]); 
      if(!dark) {
         var dx = s2[0] - s1[0], dy = s2[1] - s1[1];
         var dist = Math.sqrt(dx*dx + dy*dy);
         var offset = ((a.id + b.id) % 10 - 4.5) * 0.2 * dist; 
         var cx = (s1[0]+s2[0])/2 - (dy/dist) * offset;
         var cy = (s1[1]+s2[1])/2 + (dx/dist) * offset;
         ctx.quadraticCurveTo(cx, cy, s2[0], s2[1]);
      } else {
         ctx.lineTo(s2[0],s2[1]); 
      }
      ctx.stroke();
    }
  }

  // dots — size = page length; colour = monochrome tones (light) or genre star colour (dark)
  var showLabels=cam.s>1.7, labelCand=[];
  for(var k=0;k<BOOKS.length;k++){
    var b=BOOKS[k], sp=w2s(b.x,b.y);
    if(sp[0]<-8||sp[0]>W+8||sp[1]<-8||sp[1]>H+8) continue;
    var active=passFilter(b), read=isRead(b.id), col=dotColor(b);
    var rad=Math.max(1.2, b.rad*cam.s);
    if(!active){ ctx.fillStyle=dark?"rgba(120,132,164,.26)":"rgba(190,194,203,.4)"; ctx.beginPath(); ctx.arc(sp[0],sp[1],Math.max(1,rad*0.7),0,7); ctx.fill(); continue; }
    if(read && state.status!=="read"){ ctx.fillStyle=hexA(col,dark?.36:.28); ctx.beginPath(); ctx.arc(sp[0],sp[1],rad*0.9,0,7); ctx.fill(); }
    else {
      ctx.fillStyle=col; ctx.beginPath(); ctx.arc(sp[0],sp[1],rad,0,7); ctx.fill();
      if(dark){ 
          ctx.fillStyle="rgba(255,250,220,1)"; ctx.beginPath(); ctx.arc(sp[0],sp[1],Math.max(1.0,rad*0.5),0,7); ctx.fill(); 
          var orbitR = Math.max(8, rad*4.2);
          ctx.strokeStyle="rgba(255,255,255,0.02)";
          ctx.lineWidth=1.0;
          ctx.beginPath(); ctx.arc(sp[0],sp[1],orbitR,0,7); ctx.stroke();
          
          if(cam.s>1.8){
              var txt = (b.a||"Unknown") + " • " + b.g + " • " + b.pg + "p";
              ctx.font="9px "+FS; ctx.fillStyle="rgba(255,255,255,0.75)";
              ctx.textAlign="center"; ctx.textBaseline="middle";
              
              var tw = ctx.measureText(txt).width;
              if(tw < orbitR * 2 * Math.PI * 0.75) {
                  var totalAngle = tw / orbitR;
                  ctx.save();
                  ctx.translate(sp[0], sp[1]);
                  var timeRot = (ts / 2000) % (Math.PI * 2);
                  ctx.rotate(-Math.PI/2 - totalAngle/2 + timeRot);
                  for (var c = 0; c < txt.length; c++) {
                      var char = txt[c];
                      var cw = ctx.measureText(char).width;
                      ctx.rotate((cw/2) / orbitR);
                      ctx.fillText(char, 0, -orbitR);
                      ctx.rotate((cw/2) / orbitR);
                  }
                  ctx.restore();
                  requestDraw();
              }
          }
      } else {
          // Vibrant light mode dot core (dark shade of genre color, 1.5x larger)
          ctx.fillStyle=darken(darken(col)); ctx.beginPath(); ctx.arc(sp[0],sp[1],Math.max(1.2,rad*0.53),0,7); ctx.fill(); 
      }
    }
    if(showLabels && labelCand.length<60) labelCand.push([b,sp,read]);
  }

  // selected + hover accents
  drawAccent(state.selected,true,ts);
  if(state.hover && state.hover!==state.selected) drawAccent(state.hover,false,ts);

  // zoomed-in book labels
  if(showLabels){
    ctx.font="600 12px "+FS; ctx.textAlign="left"; ctx.textBaseline="middle";
    var lblBg=dark?"rgba(8,10,20,.6)":"rgba(255,255,255,.85)";
    var lblInk=dark?"#dfe4f2":"#3a4150", lblRead=dark?"#7f8aa6":"#9aa1ad";
    for(var i=0;i<labelCand.length;i++){
      var b=labelCand[i][0], sp=labelCand[i][1];
      var tx=sp[0]+Math.max(3,b.rad*cam.s)+4, ty=sp[1], txt=trunc(b.t,26);
      ctx.fillStyle=lblBg;
      var w=ctx.measureText(txt).width;
      ctx.fillRect(tx-2,ty-8,w+4,16);
      ctx.fillStyle=labelCand[i][2]?lblRead:lblInk; ctx.fillText(txt,tx,ty);
    }
  }

  // genre region labels — hidden at overview; fade in as you zoom, or on hover
  var zoomA = cam.s<=0.42?0 : cam.s>=0.78?1 : (cam.s-0.42)/0.36;
  ctx.textAlign="center"; ctx.textBaseline="middle";
  var drawnBoxes=[];
  // ordered: hovered genre first (never skipped), then large clusters
  var order=genres.slice();
  if(state.hoverGenre){ order=order.filter(function(x){return x!==state.hoverGenre;}); order.unshift(state.hoverGenre); }
  for(var oi=0;oi<order.length;oi++){
    var g=order[oi]; if(state.hidden[g]) continue;
    var isHover=(g===state.hoverGenre);
    var a=isHover?1:zoomA;
    if(a<=0.02) continue;
    var m=genreMeta[g], sp=w2s(m.cx,m.cy);
    var fs=Math.max(11,Math.min(27,m.r*cam.s*0.085));
    if(isHover) fs=Math.max(fs,13);
    if(sp[0]<-220||sp[0]>W+220||sp[1]<-120||sp[1]>H+120) continue;
    ctx.font="700 "+fs+"px "+FS;
    var tw=ctx.measureText(g).width, th=fs;
    var bx=sp[0]-tw/2-3, by=sp[1]-th/2-2, bw=tw+6, bh=th+4, clash=false;
    for(var d=0;d<drawnBoxes.length;d++){ var o=drawnBoxes[d]; if(bx<o[0]+o[2]&&bx+bw>o[0]&&by<o[1]+o[3]&&by+bh>o[1]){ clash=true; break; } }
    if(clash && !isHover) continue;
    drawnBoxes.push([bx,by,bw,bh]);
    ctx.fillStyle=(dark?"rgba(4,6,14,":"rgba(255,255,255,")+(0.72*a)+")"; ctx.fillText(g,sp[0]+0.6,sp[1]+0.6);
    var lf=(g==="Manga"||g==="Comics & Graphic Novels")?0.26:0.42;   // warm labels need extra darkening on white
    ctx.fillStyle=hexA(dark?lightenForDarkModeText(m.color):darken(m.color,lf),(isHover?0.98:0.9)*a); ctx.fillText(g,sp[0],sp[1]);
  }
}

function lightenForDarkModeText(hex){ if(!hex) return "#ffffff"; var h=hex.replace("#","");var r=parseInt(h.substr(0,2),16),g=parseInt(h.substr(2,2),16),b=parseInt(h.substr(4,2),16); if(r+g+b<300){ r=Math.min(255,r+120);g=Math.min(255,g+120);b=Math.min(255,b+120); } return "#"+[r,g,b].map(function(v){return ("0"+v.toString(16)).slice(-2);}).join(""); }
function darken(hex,f){ f=f||0.62; if(!hex) return "#000000"; var h=hex.replace("#","");var r=parseInt(h.substr(0,2),16),g=parseInt(h.substr(2,2),16),b=parseInt(h.substr(4,2),16); r=Math.round(r*f);g=Math.round(g*f);b=Math.round(b*f); return "#"+[r,g,b].map(function(v){return ("0"+v.toString(16)).slice(-2);}).join(""); }
function drawAccent(id,isSel,ts){
  if(id==null) return; var b=byId[id]; if(!b) return;
  var sp=w2s(b.x,b.y), rad=Math.max(3.5,b.rad*cam.s);
  ctx.fillStyle=b.color; ctx.beginPath(); ctx.arc(sp[0],sp[1],rad,0,7); ctx.fill();
  ctx.strokeStyle=isSel?(theme==="dark"?"#ffffff":"#1e2330"):(theme==="dark"?"#aab4ff":"#5b57e0"); ctx.lineWidth=isSel?2.4:1.8;
  var pulse=isSel?(rad+5+2*Math.sin((ts||0)/300)):(rad+4);
  ctx.beginPath(); ctx.arc(sp[0],sp[1],pulse,0,7); ctx.stroke();
  if(isSel){
    // dark pill label
    ctx.font="700 12px "+FS; var txt=trunc(b.t,30), w=ctx.measureText(txt).width;
    var px=sp[0]-w/2-9, py=sp[1]-pulse-26, pw=w+18, ph=20;
    roundRect(px,py,pw,ph,7); ctx.fillStyle="#1e2330"; ctx.fill();
    ctx.fillStyle="#fff"; ctx.textAlign="left"; ctx.textBaseline="middle"; ctx.fillText(txt,px+9,py+ph/2);
    // pointer
    ctx.beginPath(); ctx.moveTo(sp[0]-5,py+ph); ctx.lineTo(sp[0]+5,py+ph); ctx.lineTo(sp[0],py+ph+6); ctx.closePath(); ctx.fillStyle="#1e2330"; ctx.fill();
    requestDraw();
  }
}
function roundRect(x,y,w,h,r){ ctx.beginPath(); ctx.moveTo(x+r,y); ctx.arcTo(x+w,y,x+w,y+h,r); ctx.arcTo(x+w,y+h,x,y+h,r); ctx.arcTo(x,y+h,x,y,r); ctx.arcTo(x,y,x+w,y,r); ctx.closePath(); }
var FS='-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif';
function trunc(s,n){ return s.length>n?s.slice(0,n-1)+"…":s; }

// ---------------- picking ----------------
function pick(px,py){
  var best=null,bestD=15*15;
  for(var k=0;k<BOOKS.length;k++){ var b=BOOKS[k]; if(!passFilter(b)) continue; var sp=w2s(b.x,b.y); var dx=sp[0]-px,dy=sp[1]-py,d=dx*dx+dy*dy; var rr=Math.max(6,b.rad*cam.s+4),thr=rr*rr; if(d<thr&&d<bestD){bestD=d;best=b;} }
  return best;
}

// ---------------- camera ----------------
var userMoved=false;
function flyTo(x,y,s,dur){ cam.fx=cam.x;cam.fy=cam.y;cam.fs=cam.s;cam.tx=x;cam.ty=y;cam.ts=s;cam.animStart=performance.now();cam.animDur=dur||700;cam.animating=true;requestDraw(); }
function fitView(dur){
  var minx=1e9,miny=1e9,maxx=-1e9,maxy=-1e9;
  BOOKS.forEach(function(b){ if(b.x<minx)minx=b.x; if(b.y<miny)miny=b.y; if(b.x>maxx)maxx=b.x; if(b.y>maxy)maxy=b.y; });
  var gx=(minx+maxx)/2, gy=(miny+maxy)/2; if(!(W>0&&H>0)) return;
  var pad=60, s=Math.min((W-pad*2)/((maxx-minx)+80),(H-pad*2)/((maxy-miny)+80))*0.98;
  if(!isFinite(s)||s<=0) s=0.15;
  if(dur){ flyTo(gx,gy,s,dur); } else { cam.x=gx; cam.y=gy; cam.s=s; requestDraw(); }
}
function focusBook(b){ state.selected=b.id; openDetail(b); userMoved=true; flyTo(b.x,b.y,Math.max(2.6,cam.s),720); }

// ---------------- input ----------------
var down=false,moved=false,lastX=0,lastY=0,downX=0,downY=0,pointers={},pinchDist=0;
function localXY(e){ return [e.clientX, e.clientY-TOP]; }
cv.addEventListener("pointerdown",function(e){ down=true;moved=false;lastX=e.clientX;lastY=e.clientY;downX=e.clientX;downY=e.clientY;cam.animating=false;cv.classList.add("grabbing");cv.setPointerCapture(e.pointerId);pointers[e.pointerId]={x:e.clientX,y:e.clientY}; });
cv.addEventListener("pointermove",function(e){
  if(pointers[e.pointerId]) pointers[e.pointerId]={x:e.clientX,y:e.clientY};
  if(Object.keys(pointers).length>=2){ pinchMove(); return; }
  var lc=localXY(e);
  if(down){ var dx=e.clientX-lastX,dy=e.clientY-lastY; if(Math.abs(e.clientX-downX)+Math.abs(e.clientY-downY)>4){moved=true;userMoved=true;} cam.x-=dx/cam.s;cam.y-=dy/cam.s;lastX=e.clientX;lastY=e.clientY;requestDraw();hideTip(); }
  else {
    var b=pick(lc[0],lc[1]); if((b?b.id:null)!==state.hover){ state.hover=b?b.id:null; requestDraw(); } if(b) showTip(b,e.clientX,e.clientY); else hideTip();
    // reveal the region label under the cursor: the hovered dot's genre, else the nearest cluster
    var hg = b ? b.g : null;
    if(!hg){ var wp=s2w(lc[0],lc[1]), bd=1e18; for(var gi=0;gi<genres.length;gi++){ var g=genres[gi]; if(state.hidden[g]) continue; var m=genreMeta[g]; var dx=(wp[0]-m.cx)/STRETCHX, dy=wp[1]-m.cy, d=dx*dx+dy*dy, rr=m.r*1.05; if(d<rr*rr && d<bd){ bd=d; hg=g; } } }
    if(hg!==state.hoverGenre){ state.hoverGenre=hg; requestDraw(); }
  }
});
cv.addEventListener("pointerleave",function(){ if(state.hoverGenre){ state.hoverGenre=null; requestDraw(); } hideTip(); });
function endPointer(e){ down=false;cv.classList.remove("grabbing");delete pointers[e.pointerId];pinchDist=0; if(!moved){ var lc=localXY(e); var b=pick(lc[0],lc[1]); if(b){state.selected=b.id;openDetail(b);requestDraw();} else closeDetail(); } }
cv.addEventListener("pointerup",endPointer);
cv.addEventListener("pointercancel",function(e){ down=false;delete pointers[e.pointerId];pinchDist=0;cv.classList.remove("grabbing"); });
function pinchMove(){ var ks=Object.keys(pointers); if(ks.length<2) return; var a=pointers[ks[0]],b=pointers[ks[1]]; var dx=a.x-b.x,dy=a.y-b.y,dist=Math.sqrt(dx*dx+dy*dy); var cx=(a.x+b.x)/2,cy=(a.y+b.y)/2-TOP; if(pinchDist) zoomAt(cx,cy,dist/pinchDist); pinchDist=dist; hideTip(); }
cv.addEventListener("wheel",function(e){ e.preventDefault(); zoomAt(e.clientX,e.clientY-TOP,Math.pow(1.0016,-e.deltaY)); hideTip(); },{passive:false});
function zoomAt(px,py,factor){ userMoved=true; var wpt=s2w(px,py); cam.s=Math.max(0.03,Math.min(16,cam.s*factor)); var np=w2s(wpt[0],wpt[1]); cam.x+=(np[0]-px)/cam.s; cam.y+=(np[1]-py)/cam.s; requestDraw(); }

// ---------------- tooltip ----------------
var tip=document.getElementById("tip");
function showTip(b,x,y){
  var rt=getRating(b.id);
  tip.innerHTML='<div class="tt">'+esc(b.t)+'</div>'+(b.a?'<div class="ta">'+esc(b.a)+'</div>':'')+
    '<div class="tg" style="color:'+b.color+'"><span class="dot" style="width:8px;height:8px;background:'+b.color+'"></span>'+esc(b.g)+'</div>'+
    (rt?'<div class="trate">'+starStr(rt)+'</div>':'')+
    '<div class="ts">'+(isRead(b.id)?"✓ Read":"○ Unread")+' · '+b.f+' · '+b.pg+'pp</div>';
  tip.style.left=x+"px"; tip.style.top=(y)+"px"; tip.style.display="block";
}
function hideTip(){ tip.style.display="none"; }
function esc(s){ return (s||"").replace(/[&<>"]/g,function(c){return{"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c];}); }

// ---------------- detail ----------------
var detail=document.getElementById("detail");
function openDetail(b){
  document.getElementById("dtitle").textContent=b.t;
  document.getElementById("dauth").textContent=b.a||"Unknown author";
  var tag=document.getElementById("dtag"); tag.style.background=hexA(b.color,.13); tag.style.color=darken(b.color);
  tag.innerHTML='<span class="dot" style="width:8px;height:8px;background:'+b.color+'"></span>'+esc(b.g);
  document.getElementById("dgenre").textContent=b.g;
  document.getElementById("dformat").textContent=b.f;
  updMark(b); renderDetailStars(b);
  var sub=document.getElementById("dbysub"), byA=document.getElementById("dbyauthor"); byA.innerHTML=""; var na=norm(b.a);
  if(na.length>=4){ var same=BOOKS.filter(function(x){return x.id!==b.id&&norm(x.a)===na;}).slice(0,6); if(same.length){ sub.style.display="block"; same.forEach(function(x){ byA.appendChild(chip(x.t,function(){focusBook(x);})); }); } else sub.style.display="none"; } else sub.style.display="none";
  var gc=document.getElementById("dgenrechips"); gc.innerHTML=""; gc.appendChild(chip("◎ Show only "+b.g,function(){soloGenre(b.g);}));
  shuffle(BOOKS.filter(function(x){return x.id!==b.id&&x.g===b.g&&!isRead(x.id);})).slice(0,5).forEach(function(x){ gc.appendChild(chip(x.t,function(){focusBook(x);})); });
  detail.classList.add("on");
}
function updMark(b){
  var st=document.getElementById("dstatus"), read=isRead(b.id);
  st.textContent=read?"✓ Read":"○ Unread"; st.style.color=read?"var(--good)":"var(--muted)";
  var mb=document.getElementById("markbtn"); if(mb) mb.style.display="none"; // read-only mode
}
document.getElementById("dclose").onclick=closeDetail;
function closeDetail(){ detail.classList.remove("on"); state.selected=null; requestDraw(); }
function chip(text,fn){ var c=document.createElement("div"); c.className="chip"; c.textContent=text; c.title=text; c.onclick=fn; return c; }
// read-only 1-5 star display in the detail panel
function renderDetailStars(b){
  var host=document.getElementById("dstars"); host.innerHTML=""; var cur=getRating(b.id);
  for(var i=1;i<=5;i++){ (function(i){
    var sp=document.createElement("span"); sp.className="star"+(i<=cur?" on":""); sp.textContent="★"; sp.title=i+(i>1?" stars":" star");
    sp.style.cursor="default"; // read-only mode
    host.appendChild(sp);
  })(i); }
  document.getElementById("dratev").textContent=cur?(cur+"/5"):"Not rated";
}
function paintStars(host,n){ var ch=host.children; for(var i=0;i<ch.length;i++) ch[i].classList.toggle("on", i<n); }
function afterRating(b){
  renderDetailStars(b); updMark(b); updProgress(); buildGenreList(); requestDraw();
  var r=getRating(b.id); toast(r?("★ Rated "+r+"/5 — "+trunc(b.t,26)):("Rating cleared — "+trunc(b.t,26)));
}
function toggleRead(b){
  if(isRead(b.id)){ delete readSet[b.id]; delete ratings[b.id]; }   // marking unread clears the rating
  else { readSet[b.id]=1; }
  saveRead(); saveRatings(); updMark(b); renderDetailStars(b); updProgress(); buildGenreList(); requestDraw();
  toast(isRead(b.id)?("✓ Marked read: "+trunc(b.t,30)):("○ Marked unread: "+trunc(b.t,30)));
}
function shuffle(a){ a=a.slice(); for(var i=a.length-1;i>0;i--){var j=Math.floor(Math.random()*(i+1)),t=a[i];a[i]=a[j];a[j]=t;} return a; }

// ---------------- drawer ----------------
var drawer=document.getElementById("drawer"), scrim=document.getElementById("scrim");
function openDrawer(){ drawer.classList.add("on"); scrim.classList.add("on"); }
function closeDrawer(){ drawer.classList.remove("on"); scrim.classList.remove("on"); document.getElementById("results").classList.remove("on"); }
document.getElementById("burger").onclick=function(){ drawer.classList.contains("on")?closeDrawer():openDrawer(); };
document.getElementById("drawerClose").onclick=closeDrawer;
scrim.onclick=closeDrawer;
document.getElementById("searchIcon").onclick=function(){ openDrawer(); setTimeout(function(){document.getElementById("search").focus();},260); };
// ---------------- theme (light / dark) ----------------
function applyTheme(){ document.body.classList.toggle("theme-dark", theme==="dark"); try{localStorage.setItem(THEME_KEY,theme);}catch(e){} requestDraw(); }
document.getElementById("themeBtn").onclick=function(){ theme=(theme==="dark")?"light":"dark"; applyTheme(); toast(theme==="dark"?"Dark mode — universe":"Light mode — Brain"); };

// ---------------- Read Today ----------------
document.getElementById("today").onclick=function(){
  var pool=BOOKS.filter(function(b){ return passFilterBase(b)&&!isRead(b.id); });
  if(!pool.length){ toast("No unread books match your filters."); return; }
  var pick=pool[Math.floor(Math.random()*pool.length)]; closeDrawer(); focusBook(pick); toast("✦ Today's pick — "+trunc(pick.t,40));
};
function passFilterBase(b){ if(state.hidden[b.g]) return false; if(state.format!=="all"&&b.f!==state.format) return false; if(state.search&&(b.t+" "+b.a).toLowerCase().indexOf(state.search)<0) return false; return true; }

// ---------------- search ----------------
var search=document.getElementById("search"), results=document.getElementById("results"), resSel=-1, resList=[];
search.addEventListener("input",function(){ state.search=search.value.trim().toLowerCase(); requestDraw(); renderResults(); });
search.addEventListener("keydown",function(e){
  if(e.key==="ArrowDown"){ resSel=Math.min(resList.length-1,resSel+1); markRes(); e.preventDefault(); }
  else if(e.key==="ArrowUp"){ resSel=Math.max(0,resSel-1); markRes(); e.preventDefault(); }
  else if(e.key==="Enter"){ var t=resList[resSel]||resList[0]; if(t){ focusBook(t); results.classList.remove("on"); closeDrawer(); } }
  else if(e.key==="Escape"){ search.value="";state.search="";results.classList.remove("on");requestDraw(); }
});
function renderResults(){
  var q=state.search; if(!q){ results.classList.remove("on"); resList=[]; return; }
  var arr=BOOKS.filter(function(b){ return (b.t+" "+b.a).toLowerCase().indexOf(q)>=0; });
  arr.sort(function(a,b){ var at=a.t.toLowerCase().indexOf(q),bt=b.t.toLowerCase().indexOf(q); return (at<0?99:at)-(bt<0?99:bt); });
  resList=arr.slice(0,8); resSel=-1;
  if(!resList.length){ results.innerHTML='<div class="res"><div class="rt">No matches</div></div>'; results.classList.add("on"); return; }
  results.innerHTML="";
  resList.forEach(function(b,i){ var d=document.createElement("div"); d.className="res"; var rt=getRating(b.id); d.innerHTML='<div class="rt">'+esc(b.t)+'</div><div class="ra">'+esc(b.a||"—")+' · <span class="rg" style="color:'+darken(b.color)+'">'+esc(b.g)+'</span> · '+(isRead(b.id)?"✓":"○")+(rt?' <span class="rr">'+starStr(rt)+'</span>':'')+'</div>'; d.onmouseenter=function(){resSel=i;markRes();}; d.onclick=function(){ focusBook(b); results.classList.remove("on"); closeDrawer(); }; results.appendChild(d); });
  results.classList.add("on");
}
function markRes(){ [].forEach.call(results.children,function(c,i){ c.classList.toggle("sel",i===resSel); }); }

// ---------------- filters ----------------
var statusSeg=document.getElementById("statusSeg");
statusSeg.addEventListener("click",function(e){ var btn=e.target.closest("button"); if(!btn) return; state.status=btn.dataset.s; [].forEach.call(statusSeg.children,function(c){c.classList.toggle("on",c===btn);}); requestDraw(); });
[["fmtAll","all"],["fmtE","eBook"],["fmtB","Bookshelf"]].forEach(function(p){ document.getElementById(p[0]).onclick=function(){ state.format=p[1]; ["fmtAll","fmtE","fmtB"].forEach(function(id){document.getElementById(id).classList.toggle("on",id===p[0]);}); requestDraw(); }; });
document.getElementById("connBtn").onclick=function(){ state.showConn=!state.showConn; this.classList.toggle("on",state.showConn); requestDraw(); };

// ---------------- rating filter (min stars) ----------------
var rfStars=document.getElementById("ratefilterStars"), rfClear=document.getElementById("rateClear");
function buildRateFilter(){
  rfStars.innerHTML="";
  for(var i=1;i<=5;i++){ (function(i){
    var sp=document.createElement("span"); sp.className="star"; sp.textContent="★"; sp.title=i+"+ stars";
    sp.onmouseenter=function(){ paintStars(rfStars,i); };
    sp.onmouseleave=function(){ paintStars(rfStars,state.minRating); };
    sp.onclick=function(){ state.minRating=(state.minRating===i)?0:i; syncRateFilter(); requestDraw(); };
    rfStars.appendChild(sp);
  })(i); }
  syncRateFilter();
}
function syncRateFilter(){ paintStars(rfStars,state.minRating); rfClear.classList.toggle("on",state.minRating>0); }
rfClear.onclick=function(){ state.minRating=0; syncRateFilter(); requestDraw(); toast("Rating filter cleared"); };

// ---------------- legend ----------------
var genresEl=document.getElementById("genres");
function buildGenreList(){
  genresEl.innerHTML="";
  genres.forEach(function(g){ var m=genreMeta[g]; var readN=BOOKS.reduce(function(acc,b){return acc+((b.g===g&&isRead(b.id))?1:0);},0);
    var it=document.createElement("div"); it.className="gitem"+(state.hidden[g]?" off":"");
    it.innerHTML='<span class="dot" style="background:'+m.color+'"></span><span class="gname">'+esc(g)+'</span><span class="gcount">'+readN+'/'+m.n+'</span>';
    it.onclick=function(){ state.hidden[g]=!state.hidden[g]; it.classList.toggle("off",state.hidden[g]); requestDraw(); };
    genresEl.appendChild(it);
  });
}
document.getElementById("genAll").onclick=function(){ state.hidden={}; buildGenreList(); requestDraw(); };
function soloGenre(g){ state.hidden={}; genres.forEach(function(x){ if(x!==g) state.hidden[x]=true; }); buildGenreList(); requestDraw(); toast("Showing only "+g); }

// ---------------- zoom buttons ----------------
document.getElementById("zin").onclick=function(){ zoomAt(W/2,H/2,1.4); };
document.getElementById("zout").onclick=function(){ zoomAt(W/2,H/2,1/1.4); };
document.getElementById("zfit").onclick=function(){ userMoved=false; fitView(700); };

// ---------------- progress / toast ----------------
function updProgress(){
  document.getElementById("pcount").textContent=Object.keys(readSet).length;
  document.getElementById("tcount").textContent=BOOKS.length;
  var ids=Object.keys(ratings), n=ids.length, sum=0; for(var i=0;i<n;i++) sum+=ratings[ids[i]]||0;
  document.getElementById("ravgWrap").innerHTML = n ? (' · <span class="ravg">★'+(sum/n).toFixed(1)+'</span> avg ('+n+')') : '';
}
var toastEl=document.getElementById("toast"), toastT=0;
function toast(msg){ toastEl.textContent=msg; toastEl.classList.add("on"); clearTimeout(toastT); toastT=setTimeout(function(){toastEl.classList.remove("on");},2600); }
setTimeout(function(){ var h=document.getElementById("hint"); if(h) h.style.opacity="0"; },7000);

// ---------------- keyboard ----------------
window.addEventListener("keydown",function(e){
  if(e.target.tagName==="INPUT"||e.target.tagName==="TEXTAREA") return;
  if(e.key==="/"){ openDrawer(); setTimeout(function(){search.focus();},260); e.preventDefault(); }
  else if(e.key==="r"||e.key==="R"){ document.getElementById("today").click(); }
  else if(e.key==="f"||e.key==="F"){ userMoved=false; fitView(600); }
  else if(e.key==="m"||e.key==="M"){ drawer.classList.contains("on")?closeDrawer():openDrawer(); }
  else if(e.key==="Escape"){ closeDrawer(); closeDetail(); }
});

// apply the Goodreads seed once (merges read + ratings into your saved state)
var SEED_KEY="avs-library-seed-v1", SEED_VERSION="goodreads-1";
(function applySeed(){
  try{ if(localStorage.getItem(SEED_KEY)===SEED_VERSION) return; }catch(e){}
  try{
    if(typeof SEED_READ!=="undefined") SEED_READ.forEach(function(id){ readSet[id]=1; });
    if(typeof SEED_RATINGS!=="undefined") Object.keys(SEED_RATINGS).forEach(function(id){ ratings[id]=+SEED_RATINGS[id]; readSet[id]=1; });
    saveRead(); saveRatings(); localStorage.setItem(SEED_KEY,SEED_VERSION);
  }catch(e){}
})();

// ---------------- init ----------------
TOP=document.getElementById("topbar").offsetHeight||56;
applyTheme();
buildLayout(); buildGenreList(); buildRateFilter(); updProgress(); resize();
(function(){ var tx=cam.x,ty=cam.y,ts=cam.s; cam.s=ts*0.72; flyTo(tx,ty,ts,1200); })();
requestDraw();
})();
</script>
</body>
</html>'''

out = (HTML.replace('__BOOKS_JSON__', data_json)
           .replace('__SEED_READ__', seed_read_json)
           .replace('__SEED_RATINGS__', seed_ratings_json))
# write the app one level up, next to Books TBR.xlsx (portable: relative to this script)
dest = os.path.join(os.path.dirname(SCRATCH), 'AVs-Library.html')
with open(dest,'w') as f: f.write(out)
print('wrote', dest, len(out), 'bytes; books:', len(slim))
