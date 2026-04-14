// ===== DATA — estratti dal foglio Excel originale (prova00__1_.xlsx / Originale.xlsm) =====
const DATA = {
  // O5:P250 — 240 specie esatte
  specie: [
    'Abies alba L. - abete bianco','Abies cephalonica Loudon - abete greco','Abies nebrodensis (Lojac.) Mattei - abete dei Nebrodi','Abies nordmanniana (S.) Spach - abete del Caucaso','Acacia dealbata Link - mimosa','Acer campestre L. - acero campestre','Acer monspessulanum L. - acero minore','Acer negundo L. - acero americano','Acer opulifolium Chaix - acero alpino','Acer platanoides L. - acero riccio','Acer pseudoplatanus L. - acero di monte','Acer rubrum L. - acero scarlatto','Acer saccharum Marsh. - acero del Canada','Aesculus hippocastanum L. - ippocastano','Aesculus pavia L. - pavia','Ailanthus altissima Swingle - ailanto','Albizzia julibrissin Durazzo - gaggia','Alnus cordata (Loisel.) - ontano napoletano','Alnus glutinosa L. - ontano nero','Alnus incana Moench. - ontano bianco','Alnus spp. - ontano','Araucaria bidwilli Hook. - araucaria','Araucaria heterophylla (excelsa) - pino di Norfolk','Araucaria cunninghamii Mudie - araucaria','Arbutus unedo L. - corbezzolo','Bauhinia purpurea L.','Betula nigra','Betula papyrifera','Betula pendula Roth. - betulla bianca','Betula pubescens Ehrh. - betulla pelosa','Brachychiton acerifolius','Brachychiton populneus','Brachychiton sp.','Biota orientalis - biota','Broussonetia papyrifera Vent. - moro da carta','Calocedrus decurrens - cedro della California','Carpinus betulus L. - carpino bianco','Carya ovata (Mill.) - noce bianco','Castanea sativa Miller - castagno','Casuarina spp. - casuarina','Catalpa bignonioides Walt. - catalpa',
    "Cedrus atlantica - cedro dell'Atlante","Cedrus atlantica var. glauca - cedro dell'Atlante","Cedrus deodara G.Don - cedro dell'Himalaia",'Cedrus libani Richard - cedro del Libano','Celtis australis L. - bagolaro','Celtis occidentalis L. - bagolaro occidentale','Ceratonia siliqua L. - carrubo','Cercidiphyllum japonicum','Cercis siliquastrum L. - albero di giuda','Chamaecyparis lawsoniana - Lawson','Chamaecyparis nootkatensis - cipresso di Nootka','Chorisia speciosa','Cinnamomum camphora - canfora','Citrus sinensis Pers. - arancio','Citrus aurantium','Citrus reticulata','Cordia francisci','Corylus avellana L. - nocciolo','Corylus colurna','Cryptomeria japonica Don - criptomeria','Cupressocyparis leylandi',
    "Cupressus arizonica - cipresso dell'Arizona",'Cupressus macrocarpa - cipresso Monterey','Cupressus sempervirens L. - cipresso','Cydonia oblonga - cotogno','Diospyros kaki L. - diospero','Diospyros virginiana','Eryobotria japonica L. - nespolo giapponese','Erythrina sp. (E. christa-galli, ecc.)','Eucaliptus globulus - eucalipto','Eucaliptus camaldulensis - eucalipto','Eucalyptus gunnii','Fagus sylvatica L. - faggio','Ficus carica L. - fico','Ficus nitida','Ficus elastica','Ficus magnolioides','Ficus microcarpa','Firmiana simplex','Fraxinus excelsior L. - frassino maggiore','Fraxinus ornus L. - orniello','Fraxinus oxycarpa - frassino meridionale','Ginkgo biloba L. - ginkgo','Gleditsia triacanthos L. - spino di Giuda','Ilex aquifolium L. - agrifoglio','Grevillea robusta',
    "Hesperocyparis glabra Bartel - cip.dell'Ariz. glabro",'Jacaranda mimosifolia','Juglans nigra L. - noce nero','Juglans regia L. - noce comune','Juniperus communis L. - ginepro comune','Juniperus oxycedrus L. - ginepro rosso','Juniperus virginiana L. - ginepro della Virginia','Koelreuteria paniculata Laxm. - koelreuteria','Laburnum anagyroides','Lagunaria patersonii','Lagerstroemia indica L. - albero San Bartolomeo','Larix decidua Miller - larice','Larix kaempferi (Lamb.) - larice giapponese','Laurus nobilis L. - alloro','Ligustrum lucidum Ait. - ligustro lucido','Liquidambar styraciflua L. - liquidambar','Liriodendron tulipifera L. - albero dei tulipani','Maclura pomifera Schneider - maclura','Magnolia denudata - magnolia yulan','Magnolia grandiflora L. - magnolia','Magnolia obovata - magnolia giapponese','Magnolia stellata - magnolia stellata','Magnolia x soulangeana Solu.Bod. - magnolia',
    'Malus floribunda - melo da fiore','Malus sylvestris - melo','Melia azedarach - albero dei rosari','Mespilus germanica - nespolo','Metasequoia glyptostroboides - metasequoia','Morus alba - gelso bianco','Morus nigra - gelso nero','Musa paradisiaca - banano','Nerium oleander - oleandro','Nyssa sylvatica','Olea europea - olivo','Ostrya carpinifolia - carpino nero','Parrotia persica - parrozia','Paulownia tomentosa - paulownia','Phillirea latifolia - fillirea','Phellodendron amurense','Phoenix canariensis - palma delle Canarie','Phoenix dactylifera - palma da datteri','Phytolacca dioica - fitolacca','Picea abies - abete rosso','Picea pungens - abete del Colorado','Pinus brutia - pino calabrese','Pinus canariensis - pino delle Canarie','Pinus cembra - pino cembro',
    "Pinus halepensis - pino d'Aleppo",'Pinus mugo - pino montano','Pinus nigra - pino nero','Pinus pinaster - pino marittimo','Pinus pinea - pino domestico',"Pinus ponderosa - pino giallo dell'ovest",'Pinus radiata - pino insigne','Pinus strobus - pino strobo','Pinus sylvestris - pino silvestre','Pinus uncinata - pino uncinato',"Pinus wallichiana - pino dell'Himalaia",'Pistacia terebinthus - terebinto','Platanus occidentalis - platano comune','Platanus orientalis - platano orientale','Platanus x acerifolia - platano ibrido','Platycladus orientalis','Populus alba - pioppo bianco','Populus canescens - pioppo gatterino','Populus nigra - pioppo nero','Populus nigra var. italica - pioppo cipressino','Populus spp. - pioppo','Populus tremula - pioppo tremolo',
    'Prunus amygdalus - mandorlo','Prunus armeniaca - albicocco','Prunus avium - ciliegio','Prunus cerasifera - mirabolano','Prunus domestica - susino','Prunus glandulosa - mandorlo da fiore','Prunus laurocerasus L. - lauroceraso','Prunus pissardi','Prunus serrulata','Prunus subhirtella','Pseudotsuga menziesii - abete di Douglas','Ptelea trifoliata - olmo di Samaria','Pterocarya fraxinifolia - noce del Caucaso','Pyrus calleriana','Pyrus pyraster - pero selvatico','Quercus cerris - cerro','Quercus frainetto - farnetto','Quercus ilex - leccio','Quercus palustris - quercia delle paludi','Quercus petraea - rovere','Quercus pubescens - roverella','Quercus robur - farnia','Quercus rubra - quercia rossa','Quercus suber - sughera','Quercus trojana',
    'Rhamnus alaternus - alaterno','Rhus typhina - sommaco','Robinia pseudoacacia - robinia','Salix alba - salice bianco','Salix babylonica - salice piangente','Salix caprea - salicone','Salix daphnoides - salice nero','Salix spp. - salice','Sequoia sempervirens - sequoia','Sequoiadendron giganteum - sequoia gigante','Schinus molle','Sophora japonica - sofora','Sorbus aria - sorbo montano','Sorbus aucaparia - sorbo degli uccellatori','Sorbus domestica - sorbo domestico','Sorbus torminalis - ciavardello','Styphnolobium japonicum','Tamarix gallica - tamerice','Taxodium disticum - cipresso calvo','Taxus baccata - tasso','Thuja occidentalis - tuia','Thuja orientalis - albero della vita','Thuja plicata - tuia gigante',
    'Tilia cordata - tiglio selvatico','Tilia platyphyllos - tiglio nostrano','Tilia sp. - tiglio','Tilia x europaea - tiglio ibrido','Tilia x vulgaris - tiglio ibrido','Trachycarpus fortunei - trachycarpus','Tsuga canadensis - tsuga canadese','Ulmus glabra - olmo montano','Ulmus maior','Ulmus minor - olmo campestre','Ulmus parvifolia','Ulmus procera','Ulmus pumila - olmo siberiano','Washingtonia filifera - washingtonia','Washingtonia robusta - washingtonia','Yucca spp.','Zelkova carpinifolia - zelkova','Zelkova spp.','Ziziphus jujuba - giuggiolo comune',
    '--- PALME ---','Brahea armata','Chamaerops humilis L. - palma nana','Jubaea chilensis','Phoenix canariensis C. - palma delle Canarie','Phoenix dactylifera L. - palma da datteri','Trachycarpus fortunei Wendl. - trachycarpus','Washingtonia filifera Wendl. - washingtonia','Washingtonia robusta - washingtonia','Yucca spp.'
  ],
  // U4:V27
  dimora: ['prato','scarpata','terrapieno','terreno coltivato','terreno incolto','buco asfalto','area di pertinenza','banchina stradale','rimboschimento','tornello','aiuola','bosco','aiuola spartitraffico','alberata stradale','parcheggio','gruppo/boschetto','giardino recente','filare arboreo','piazza','cimitero','parco recente','giardino storico','parco storico','marker'],
  // R4:S10
  stadio: ['plantula','giovane pianta','albero giovane','albero adulto','albero maturo','albero senescente','albero veterano'],
  // R11:S15 — palma veterana inclusa
  stadioPalme: ['palma giovane','palma adulta','palma senescente','palma veterana'],
  stadioPop: ['piantagione'],
  // X4:Y12
  posSociale: ['oppressa','dominata','intermedia','codominante','dominante margine','dominante interna','predominante','libera (p giovane)','isolata'],
  // AA4:AB11 — 'centro città' con accento
  localiz: ['aree rurali','aree rurali urbaniz.','aree industriali','periferia recente','periferia antica','luoghi villeggiatura','centro città','centro storico'],
  // AD4:AE9
  vincoli: ['---','> soglia comunale','paesaggistico','dichiaraz. di rilevanza','storico-architettonico','monumentale'],
  // AD12:AE16
  coerFito: ['A: optimum','B: elevata','C: buona','D: scadente','E: incoerente'],
  // I4:I22 — 19 voci con 'pedoni' e 'veicoli'
  conflitti: ['assenti','asfalto','pavimentazione','cavi storrenaei','cavi aerei','cordoli','fabbricati','lampioni','manufatti','marciapiedi','muri','oggetti inglobati','cartellonistica','veg. erbacea','arbusti','lianose','altri alberi','pedoni','veicoli'],
  // K4:K50 — 47 agenti carie
  agentiCarie: [
    'non segni cariogeni','carie non determin.','Abortiporus biennis','Agrocybe aegerita','Armillaria sp.','Cerrena unicolor','Collybia fusipes','Chondrostereum p.','Coniophora puteana','Daedalea quercina','Daldinia concentrica','Fistulina epatica','Fomes fomentarius','Fomitopsis hiberica','Fomitopsis pinicola',
    'Ganoderma applanatum','Ganoderma australe','Ganoderma adspersum','Ganoderma lipsiense','Ganoderma lucidum','Ganoderma resinaceum','Grifola frondosa','Heterobasidion annosum','Inonotus cuticularis','Inonotus hispidus','Inonotus dryadeus','Inonotus rickii','Inonotus tamaricis','Laetiporus sulphureus','Meripilus giganteus',
    'Oxyporus latemarginatus','Oxyporus populinus','Perenniporia fraxinea','Phaeolus schweinitzii','Phellinus igniarius','Phellinus pini','Phellinus pomaceus','Phellinus punctatus','Phellinus robustus','Phellinus torulosus','Pholiota spp.','Polyporus squamosus','Rigidoporus ulmarius','Spongipellis pachiodon','Spongipellis pumeus','Thielaviopsis paradoxa','Trametes gibbosa'
  ],
  // M4:M16 — altri patogeni (senza duplicati)
  altriPatogeni: ['non segni fitopatie','patog. non determinati','Processionaria','Tomicus','Punteruolo rosso','Paisandisia','Grafiosi','Tortrice','Cancro corticale','Cancro colorato','Scolitidi','Cerambicidi','altro'],
  // E4:E13 — monitoraggio (senza 'non definito' se non presente)
  monitoraggio: ['6 mesi','1 anno','2 anni','3 anni','4 anni','5 anni','6 anni','10 anni','subito','stagionale','non definito'],
  // G4:G13
  urgenza: ['subito','< 6 mesi','1 anno','2 anni','3 anni','4 anni','al bisogno','opzionale','differibile','consigliato','non definito'],
  // AT35:AU45 — 11 voci condizioni salute (CORRETTE con voce 4 completa)
  condizSalute: [
    'Condizioni vegetative e fitosanitarie ottimali. Albero integro',
    'Condizioni vegetative e/o fitosanitarie ottimali. Albero lievemente alterato nella struttura',
    'Condizioni vegetative e/o fitosanitarie buone o comunque non tali da condizionare la salute e la vigoria',
    'Condizioni vegetative e/o fitosanitarie buone o comunque non tali da condizionare la salute e la vigoria. Albero strutturalmente alterato',
    "Condizioni vegetative e/o fitosanitarie mediocri, che limitano l'efficienza funzionale. Salute e/o vigoria ridotte",
    'Condizioni vegetative e/o fitosanitarie mediocri. Albero strutturalmente alterato',
    "Condizioni vegetative e/o fitosanitarie scadenti, che ne condizionano la salute e l'aspettativa di vita",
    'Condizioni vegetative e/o fitosanitarie scadenti. Albero molto alterato strutturalmente',
    'Condizioni vegetative e/o fitosanitarie pessime',
    'Condizioni vegetative e/o fitosanitarie pessime. Albero fortemente deperiente, strutturalmente molto alterato',
    'Albero morto in piedi'
  ],
  // BF4:BF11 — tipo bersaglio TUTTO minuscolo come da originale
  tipoBersaglio: ['proprietà','occupazione','pedoni/ciclisti','traffico 30 km/h','traffico 50 km/h','traffico 70 km/h','traffico 90 km/h','traffico 110 km/h'],
  // Y5:Y8 — aspetto organo ORD
  aspettoOrgano: ['ps','s','ms'],
  // ORD AA2:AA45 — zolla/radici
  radici: ['compattazione','competizione radicale','erosione','pavimentazione','ristagno idrico','rottura suolo (fessure)','scavi o trincee','sollevamento manufatti','platea radicale','ricarico zolla','sollevamento zolla','sprofondamento zolla','zolla asimmetrica','zolla danneggiata','zolla limitata / ridotta','zolla/suolo separati','esploratrici esposte','polloni radicali','noduli in superficie','radici affioranti','radici avventizie','radici avvolgenti','radici decorticate','radici esposte','radici fini visibili','radici migranti','radici non visibili','radici scoperte','radici strozzanti','radici tagliate','carie/marciume radicale','carpofori','ferite aperte','ferite richiuse','Inclusione/contatto','legno disfunzionale','necrosi radicali','non valutabile'],
  // ORD AC2:AC62 — colletto
  collettoDiag: ["colletto a contrafforti","colletto a gradino","colletto a imbuto","colletto a palafitta","colletto asimmetrico","colletto azzampato","colletto cilindrico","colletto cordonato","colletto normale","colletto ovalizzato","colletto sepolto","collo bottiglia","affondam in comprex","cordoni xilematici","cordoni non più attivi","corna di ariete","cucitura","piatto morfologico","pieghe a molla","pieghe fisarmonica","polloni","ponte xilematico su cretto","rialzamento in traz.","rigonf. anello","spiralizzazione fibre","succhioni","abrasioni","apertura","aree funzionali inattive","bruciatura","bucature","carie basale","carpofori","cavità","cretti elicoidali","cretti orizz","cretti vert.","danni alburno/cortic","delaminazione","depressione","essudati/resinaz","ferite richiuse","fessure","gibbosità","incavo","inclusione/contatto","legno da ferita attivo","legno disfunzionale","legno reazione insuff","neoplasie/cancri","penetrazione","alteraz da animali","batteriosi","insetti/danni insetti","parassiti/epifite","non valutabile"],
  // ORD AE2:AE70 — fusto
  fustoDiag: ["fusto regolare","fusto bi (tri)forcato","fusto policormico","fusto asimmetrico","fusto costolato","fusto sciabolato","perdita leader","aree funzionali inattive","carico in punta","cordoni xilematici","corna di ariete","cucitura","inclinaz. Arcuatura","inclinaz. Sciabolatura","inclinazione lineare","piega a S","pieghe fisarmonica","ponte xilematico su cretto","reiterati ageotropi","reiterati ortotropi","reiterati plagiotropi","rigonf. anello","sinuosità","snellezza","spiralizzazione fibre","strozzatura","abrasioni","apertura","bombatura","bruciatura","bucature","carie","carpofori","cavità","corteccia inclusa","cretti all'inserz.","cretti elicoidali","cretti orizz","cretti vert.","danni alburno/cortic","delaminazione","deprex all'inserz.","deviazione","essudati/resinaz","ferite radiali","ferite richiuse","fessure","Getti epicormici","gibbosità","incavo","inclusione/anastomosi","inserzione stretta","legno da ferita attivo","legno disfunzionale","legno reazione insuff","monconi","mozziconi","neoplasie/cancri","scosciature","alteraz da animali","batteriosi","insetti/danni insetti","parassiti/epifite","non valutabile"],
  // ORD AG2:AG64 — castello
  castelloDiag: ["cast. con sbrancamento","castello a cespuglio","castello a torrioni","castello allargato","castello assente","castello regolare","codominanza","aree funzionali inattive","corna di ariete","cucitura","pieghe fisarmonica","reiterati ageotropi","reiterati ortotropi","reiterati plagiotropi","rigonf. anello","spiralizzazione fibre","ponte xilematico su cretto","colonne xilematiche","abrasioni","apertura","bombatura","bruciatura","bucature","carie","carpofori","cavità","corteccia inclusa","cretti all'inserz.","cretti elicoidali","cretti orizz","cretti vert.","danni alburno/cortic","delaminazione","ferite radiali","ferite richiuse","fessure","gibbosità","incavo","inclusione/anastomosi","inserzione debole","inserzione stretta","legno da ferita attivo","legno disfunzionale","legno reazione insuff","monconi","mozziconi","naso bulldog","naso pinocchio","neoplasie/cancri","resinaz/essudati","rigonfiamenti","scosciature","subsidenza","alteraz da animali","batteriosi","insetti/danni insetti","parassiti/epifite","non valutabile"],
  // ORD AI2:AI75 — ramificazione
  ramificazDiag: ["conforme all'habitus","rami candel/pennello","rami sinuosi/arcuati","ramificaz insufficiente","ramificaz lacunosa","aree funzionali inattive","branca deperiente","capitozzatura","cedimenti pregressi","corna di ariete","cucitura","disseccamenti branche","disseccamenti rami","disseccamenti ramuli","end loading/codaleone","fusticini epicormici","pali su branche","pali su fusto","pieghe fisarmonica","pipe","rami patenti","reiterati ageotropi","reiterati ortotropi","reiterati plagiotropi","ricacci","snellezza, ecc. L/D","sostituti","torsioni/pieghe","unità minimali","zampe di cane","abrasioni","apertura","banderuole","bombatura","bruciatura","bucature","carie","carpofori","cavità","cavità a bicchiere","corteccia inclusa","cretti all'inserz.","cretti elicoidali","cretti orizz","cretti vert.","danni alburno/cortic","delaminazione (t. sventura)","ferite radiali","ferite richiuse","fessure","incavo","inclusione/anastomosi","inser debole","inserz orizzont","inserzione stretta","legno da ferita attivo","legno disfunzionale","legno reazione insuff","monconi","mozziconi","neoplasie/cancri","resinaz/essudati","scosciature","alteraz da animali","batteriosi","insetti/danni insetti","parassiti/epifite","non valutabile"],
  // ORD AK2:AK37 — chioma
  chiomaDiag: ["conforme all'habitus","chioma a bandiera","chioma a pennello","chioma asimmetrica","chioma diradata","chioma esposta","chioma lacunosa","chioma separata","chioma spalcata","chioma ridotta","chioma sana","chioma stressata","chioma resiliente","ripiegamento della chioma","discesa della cima","deperim. irreversibile","assenza leader","compress laterale","compress superior","conflitti in chioma","discontinuità corona","LCR ridotto","trasparenza","conflitto sostituti (false cime)","nido di cicogna","microfillia","alteraz cromatica","alteraz morfologica","galle","insetti o loro nidi","necrosi fogliari","non valutabile"],
  // Prescrizioni colturali C25:C70 (senza intestazioni MAIUSCOLE)
  prescColturali: [
    'da definire dopo la valutazione','abbattimento','abbattimento con sostituzione nel sito di vegetazione','abbattimento senza sostituzione nel sito di vegetazione',
    'abbattimento nel medio periodo in un ottica di rinnovazione colturale','abbattimento non urgente per mancanza di idoneità, limitata aspettativa di vita, ecc.',
    '1A) potatura di formazione in alberi giovani (in conformità con EAS)','1D) impostazione della forma obbligata su alberi giovani (in conformità con EAS)',
    '2A) potatura di mantenimento in alberi adulti (in conformità con EAS)','2B) riduzione laterale in alberi adulti  (in conformità con EAS)',
    '2C) potatura di mantenimento della  forma condizionata in alberi adulti   (in conformità con EAS)',
    '2D) potatura di mantenimento della forma obbligata in alberi adulti  (in conformità con EAS)',
    '3A) potatura di mantenimento in alberi maturi (in conformità con EAS)','3B) riduzione laterale in alberi maturi (in conformità con EAS)',
    '3C) riduzione in altezza in alberi maturi (in conformità con EAS)','4) potatura di alberi veterani  (in conformità con EAS)',
    '5) restauro della forma naturaliforme  (in conformità con EAS)','6) restauro di forma artificiale (in conformità con EAS)',
    '7) rimonda del seccume (in conformità con EAS)','8) potatura di stabilizzazione o pre-abbattimento (in conformità con EAS)',
    'Rimozione /riduzione ramificazioni compromesse','trattamento fitosanitario','trattamento nutrizionale','diserbo',
    'supporto','ancoraggio','sostegno','progettazione e realizzazione di consolidamento','puntellamento',
    'miglioramento del terreno','arieggiamento del terreno /decompattamento','drenaggio','mulching',
    "riordino del sito di radicazione","protezione dell'albero","costruzione parafulmine",
    'Albero di scarso valore ornamentale. Abbattimento consigliato non urgente. Cure colturali antieconomiche e non risolutive.',
    'Albero NON idoneo al contesto. Abbattimento prescritto non urgente.',
    'Albero compromesso fisiologicamente (es. dominato, sovrannumerario o con poco spazio a disposizione). Abbattimento non urgente per motivi colturali.',
    'Albero morto in piedi. Abbattimento non urgente.',
    'Condizioni vegetative e/o fitosanitarie pessime. Albero compromesso ma abbattimento non urgente',
    'Condizioni vegetative e/o fitosanitarie pessime. Albero compromesso e pericoloso. Abbattimento urgente',
    'Albero morto in piedi. Porzioni in disfacimento. Abbattimento urgente.',
    'Albero da sottoporre a interventi di trattamento selvicolturale.','altro da specificare'
  ],
  // Prescrizioni valutative E16:E30 (senza header e senza duplicati)
  prescValutative: [
    'controllo inclinazione','TRIAGE','valutaz ordinaria','valutazione in quota',
    'valutazione avanzata strumentale','V. avanzata + penetrom. al colletto','V. avanzata + penetrom. sul tronco',
    'V. avanzata + penetrom. in quota','V. avanzata + tomografia al colletto','V. avanzata + tomografia sul tronco',
    'V. avanzata + tomografia in quota','V. avanzata + test trazione','Analisi radicale con scavo','altro da specificare'
  ],
  // Mitigazione bersaglio G16:G30
  mitigazBersaglio: [
    'da definire dopo la valutazione','foglio informativo','pannello informativo','delimitazione percorsi',
    'cartello di pericolo','divieto di sosta temporaneo','divieto di sosta permanente','divieto di passaggio',
    'delimitazione area a rischio','transennatura area a rischio','recinzione area a rischio',
    'spostamento del bersaglio','altro da specificare'
  ],
  // Grado pericolo TRIAGE B17:C23
  gradoPericoloTriage: {
    1:'Segni sintomi e difetti molto gravi, connessi strettamente con un possibile cedimento immediato',
    2:'Segni sintomi e difetti gravi, connessi strettamente con un possibile cedimento imminente',
    3:'Segni sintomi e difetti significativi, connessi strett. con un possibile cedimento in cond. critiche',
    4:'Segni sintomi e difetti incerti, connessi occasionalmente con un possibile cedimento',
    5:'Segni sintomi e difetti moderati, non chiaramente connessi con un possibile cedimento',
    6:'Segni sintomi e difetti modesti, non connessi con un possibile cedimento',
    7:'Segni sintomi e difetti trascurabili o sostanzialmente assenti'
  },
  // Grado pericolo ORD B4:C14
  gradoPericoloOrd: {
    1:"Segni, sintomi e difetti gravissimi - Probabilità di cedimento nell'anno 1/1-1/5",
    2:"Segni, sintomi e difetti gravi - Probabilità di cedimento nell'anno 1/5-1/50",
    3:"Segni, sintomi e difetti significativi - Probabilità di cedimento nell'anno 1/50-1/500",
    4:"Segni, sintomi e difetti incerti - Probabilità di cedimento nell'anno 1/500-1/5K",
    5:"Segni, sintomi e difetti moderati - Probabilità di cedimento nell'anno 1/5K-1/50K",
    6:"Segni, sintomi e difetti bassi - Probabilità di cedimento nell'anno 1/50K-1/500k",
    7:"Segni, sintomi e difetti trascurabili - Probabilità di cedimento nell'anno 1/500K-1/1M"
  },
  // Triage Rischio (R) — AO24:AP41 (classi 1-6)
  triageRischio: {
    1:"EMERGENZA - Risoluzione dell'emergenza",
    2:"EMERGENZA - Valutazione immediata - risoluzione dell'emergenza",
    3:'URGENZA - Valutazione urgente',
    4:'MODERATA - valutazione opportuna entro breve tempo',
    5:'BASSA - valutazione opportuna ma non urgente',
    6:'DIFFERIBILE - valutazione procrastinabile'
  }
};


const RISK_TABLE = {
  '1-1':1,'1-2':1,'1-3':1,'1-4':1,'1-5':2,'1-6':3,'1-7':3,
  '2-1':1,'2-2':1,'2-3':1,'2-4':2,'2-5':3,'2-6':4,'2-7':4,
  '3-1':1,'3-2':1,'3-3':2,'3-4':3,'3-5':4,'3-6':5,'3-7':5,
  '4-1':1,'4-2':2,'4-3':3,'4-4':4,'4-5':5,'4-6':5,'4-7':6,
  '5-1':2,'5-2':3,'5-3':4,'5-4':5,'5-5':5,'5-6':6,'5-7':6,
  '6-1':3,'6-2':4,'6-3':5,'6-4':5,'6-5':6,'6-6':6,'6-7':7,
  '7-1':3,'7-2':4,'7-3':5,'7-4':6,'7-5':6,'7-6':7,'7-7':7,
};
const RISK_LABELS = {
  1:{label:'Rischio inaccettabile',sub:'Abbattimento o eliminazione del rischio',cls:'risk-1',color:'#b71c1c',bg:'#ffebee'},
  2:{label:'Rischio inaccettabile',sub:'Abbattimento o eliminazione del rischio',cls:'risk-1',color:'#b71c1c',bg:'#ffebee'},
  3:{label:'Rischio tollerabile per accordo',sub:'Se il valore e molto elevato',cls:'risk-2',color:'#bf360c',bg:'#fff3e0'},
  4:{label:'Rischio tollerabile ALARP',sub:'Valutare costi/benefici del controllo',cls:'risk-3',color:'#f57f17',bg:'#fff9c4'},
  5:{label:'Rischio tollerabile ALARP',sub:'Valutare costi/benefici del controllo',cls:'risk-3',color:'#f57f17',bg:'#fff9c4'},
  6:{label:'Rischio tollerabile',sub:'Individuare interventi affinche rimanga ALARP',cls:'risk-4',color:'#33691e',bg:'#f1f8e9'},
  7:{label:'Rischio largamente accettabile',sub:'Nessun intervento urgente necessario',cls:'risk-5',color:'#1b5e20',bg:'#e8f5e9'},
};

// ===== FORMULE DAL FOGLIO EXCEL ORIGINALE =====
// H9 = E9*PI()  — Circonferenza (cm) da Diametro tronco (cm)
function calcCirc(d_tr) {
  const v = parseFloat(d_tr);
  return isNaN(v)||v<=0 ? null : v * Math.PI;
}

// G10 = PI()/4*(E9/100)^2*C9*0.9*900  — Biomassa (kg)
// E9=D_tronco(cm), C9=H_albero(m)
function calcBio(d_tr, h) {
  const d=parseFloat(d_tr), hv=parseFloat(h);
  if(isNaN(d)||isNaN(hv)||d<=0||hv<=0) return null;
  return (Math.PI/4) * Math.pow(d/100, 2) * hv * 0.9 * 900;
}

// J10 = (G10/VLOOKUP(stadio,BK:BM,3))*VLOOKUP(localiz,BO:BX,col)
// CO2 approssimata: bio * coeff_medio_stadio * coeff_medio_loc
// Coefficienti medi da tabella A: stadio medio=1.0, loc media ~0.55*0.9
function calcCO2(bio) {
  if(bio===null) return null;
  return bio * 0.55 * 0.9; // coeff medio stadio adulto × loc media
}

// M10 = CO2/44.01*31.999*0.9  — O2 (kg/anno) — dalla formula Excel esatta
function calcO2(co2) {
  if(co2===null) return null;
  return (co2 / 44.01) * 31.999 * 0.9;
}

// O10 = (G10*0.2)/VLOOKUP(stadio)*VLOOKUP(localiz)
function calcIA(bio) {
  if(bio===null) return null;
  return bio * 0.2 * 0.55 * 0.9;
}

// J18/J20 = (PI()/4*((H9/100)/PI())^2*C9*0.75*900) * SQRT(3*9.8*(C9-T9))
// H9=Circ(cm), C9=H(m), T9=H_bersaglio(m)
// Nota: (H9/100)/PI() = D_tronco/100 (torna al diametro)
function calcImpulsoAlbero(circ_cm, h, h_bersaglio) {
  const c=parseFloat(circ_cm), hv=parseFloat(h), hb=parseFloat(h_bersaglio);
  if(isNaN(c)||isNaN(hv)||isNaN(hb)||c<=0||hv<=hb) return null;
  const d_m = (c/100) / Math.PI;          // D_tronco in metri = Circ/(100*PI)
  const massa = (Math.PI/4) * Math.pow(d_m, 2) * hv * 0.75 * 900;
  const vel   = Math.sqrt(3 * 9.8 * (hv - hb));
  return massa * vel;
}

// S18/S20 = (PI()/4*(M9/100)^2*O9*0.75*900) * SQRT(2*9.8*(R9-T9))
// M9=D_branca(cm), O9=L_branca(m), R9=H_branca(m), T9=H_bersaglio(m)
function calcImpulsoRamo(d_br, l_br, h_br, h_bersaglio) {
  const db=parseFloat(d_br), lb=parseFloat(l_br), hbr=parseFloat(h_br), hb=parseFloat(h_bersaglio);
  if(isNaN(db)||isNaN(lb)||isNaN(hbr)||isNaN(hb)||db<=0||hbr<=hb) return null;
  const massa = (Math.PI/4) * Math.pow(db/100, 2) * lb * 0.75 * 900;
  const vel   = Math.sqrt(2 * 9.8 * (hbr - hb));
  return massa * vel;
}

// L18/L20/L33/L52 — Classe impulso: IF(I<5,7,IF(I<10,6,...))
function classeImpulso(v) {
  if(v===null||v===undefined||isNaN(v)) return '—';
  if(v<5) return 7; if(v<10) return 6; if(v<50) return 5;
  if(v<100) return 4; if(v<1000) return 3; if(v<10000) return 2;
  return 1;
}

// Calcolo Triage Rischio (R) da probabilità cedimento
// Tabella AA26:AB43 in TRG-P: prob → classe 1-6
// La prob viene dal lookup bersaglio/impulso; qui usiamo RISK_TABLE (pericolo × bersaglio) → classe ORD
// Per TRG il rischio ha una scala diversa: 1=EMERGENZA ... 6=DIFFERIBILE
function calcTriageRischio(pericolo, classeBersaglio) {
  if(!pericolo||!classeBersaglio) return null;
  const rc = RISK_TABLE[pericolo+'-'+classeBersaglio];
  return rc || null;
}

// Helper descrizione grado pericolo per hint UI
function hintPericolo(v, type='trg') {
  if(!v) return '';
  const table = type==='ord' ? DATA.gradoPericoloOrd : DATA.gradoPericoloTriage;
  return table[parseInt(v)] || '';
}

function fmtN(v,dec=1){
  if(v===null||v===undefined||isNaN(v)) return '—';
  return parseFloat(v).toFixed(dec);
}

// ===== VALORE ECOLOGICO E COMPLESSIVO =====
// Formula da Excel: D10 = Bio*0.55 + CO2*1 + O2*5 + IA*10
function calcValoreEcologico(bio,co2,o2,ia){
  if(bio===null||co2===null||o2===null||ia===null) return null;
  return bio*0.55+co2*1+o2*5+ia*10;
}
function calcValoreComplessivo(ve,bio,condiz){
  if(ve===null||bio===null) return null;
  const idx=condiz?DATA.condizSalute.indexOf(condiz):0;
  const dep=idx>=0?idx*10:0;
  return ve+Math.max(0,bio*2.5*(1-dep/100));
}
function fmtEuro(v){
  if(v===null||v===undefined||isNaN(v)) return '—';
  return '€ '+Math.round(v).toLocaleString('it-IT');
}

// ===== STATE =====
let currentTab='home';
let formData={trgp:{},trgs:{},ord:{}};
let savedForms=JSON.parse(localStorage.getItem('arete_forms')||'[]');
let editingId=null;
let archiveFilter='all';

function sel(id){return document.getElementById(id);}
function today(){return new Date().toISOString().split('T')[0];}
function uid(){return Date.now().toString(36)+Math.random().toString(36).substr(2,5);}

function showToast(msg,type=''){
  const t=sel('toast');
  t.textContent=msg; t.className='toast'+(type?' '+type:''); t.classList.add('show');
  setTimeout(()=>t.classList.remove('show'),2600);
}
function makeSelect(name,options,placeholder,value){
  const opts=options.map(o=>`<option value="${o}" ${o===value?'selected':''}>${o}</option>`).join('');
  return `<select class="form-select" name="${name}" onchange="onFieldChange('${name}',this.value)"><option value="">${placeholder||'Seleziona...'}</option>${opts}</select>`;
}
function makeInput(name,type,placeholder,value,extra){
  return `<input class="form-input" type="${type||'text'}" name="${name}" placeholder="${placeholder||''}" value="${value||''}" ${extra||''} oninput="onFieldChange('${name}',this.value)">`;
}

// ===== FORMULE LIVE — sezione misure con valori calcolati =====
function buildMisureSection(d, ft, hasEco) {
  const circ = calcCirc(d.d_tronco);
  const bio  = calcBio(d.d_tronco,d.h_albero);
  const co2  = calcCO2(bio);
  const o2   = calcO2(co2);
  const ia   = calcIA(bio);
  const imp_alb = calcImpulsoAlbero(circ,d.h_albero,d.h_bersaglio);
  const imp_ram = calcImpulsoRamo(d.d_branca,d.l_branca,d.h_branca,d.h_bersaglio);
  const cls_alb = classeImpulso(imp_alb);
  const cls_ram = classeImpulso(imp_ram);

  const ve=calcValoreEcologico(bio,co2,o2,ia);
  const vc=calcValoreComplessivo(ve,bio,d.condiz_salute);
  const ecoBar = hasEco ? `
    <div class="eco-bar">
      <div class="eco-item"><span class="eco-label">Bio <small>(kg)</small></span><span class="eco-val" id="bio-${ft}">${bio?fmtN(bio,0):'—'}</span></div>
      <div class="eco-item"><span class="eco-label">CO₂ <small>(kg/a)</small></span><span class="eco-val" id="co2-${ft}">${co2?fmtN(co2,0):'—'}</span></div>
      <div class="eco-item"><span class="eco-label">O₂ <small>(kg/a)</small></span><span class="eco-val" id="o2-${ft}">${o2?fmtN(o2,0):'—'}</span></div>
      <div class="eco-item"><span class="eco-label">I.A. <small>(kg/a)</small></span><span class="eco-val" id="ia-${ft}">${ia?fmtN(ia,0):'—'}</span></div>
    </div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:8px">
      <div style="background:#e8f5e9;border:1.5px solid #81c784;border-radius:10px;padding:10px 14px">
        <div style="font-size:11px;font-weight:600;color:#2e7d32;text-transform:uppercase;margin-bottom:4px">Valore ecologico</div>
        <div style="font-size:18px;font-weight:700;color:#1b5e20;font-family:'DM Mono',monospace" id="val-ecol-${ft}">${fmtEuro(ve)}</div>
        <div style="font-size:10px;color:#4caf50;margin-top:2px">Bio×0.55+CO₂+O₂×5+I.A.×10</div>
      </div>
      <div style="background:#e3f2fd;border:1.5px solid #90caf9;border-radius:10px;padding:10px 14px">
        <div style="font-size:11px;font-weight:600;color:#1565c0;text-transform:uppercase;margin-bottom:4px">Valore complessivo</div>
        <div style="font-size:18px;font-weight:700;color:#0d47a1;font-family:'DM Mono',monospace" id="val-comp-${ft}">${fmtEuro(vc)}</div>
        <div style="font-size:10px;color:#64b5f6;margin-top:2px">Val.ecol.+Val.ornamentale</div>
      </div>
    </div>` : '';

  return `<div class="form-section">
    <div class="form-section-header"><span class="section-icon">📐</span>Misure Dendrologiche
      <span style="font-size:10px;color:var(--sage);font-weight:400;margin-left:8px">🔁 valori calcolati automaticamente</span>
    </div>
    <div class="form-grid">
      <div class="form-row cols-3">
        <div class="form-field"><label class="form-label">H <span class="unit">(m)</span></label>${makeInput('h_albero','number','Alt.',d.h_albero,'step="0.5" min="0"')}</div>
        <div class="form-field"><label class="form-label">D tr <span class="unit">(cm)</span></label>${makeInput('d_tronco','number','Ø tr.',d.d_tronco,'step="0.5" min="0"')}</div>
        <div class="form-field"><label class="form-label">Circ <span class="unit">(cm)</span> 🔁</label><div class="computed-field" id="circ-${ft}">${circ?fmtN(circ,1):'—'}</div></div>
      </div>
      <div class="form-row cols-3">
        <div class="form-field"><label class="form-label">D ch <span class="unit">(m)</span></label>${makeInput('d_chioma','number','Ø ch.',d.d_chioma,'step="0.5" min="0"')}</div>
        <div class="form-field"><label class="form-label">D br <span class="unit">(cm)</span></label>${makeInput('d_branca','number','Ø br.',d.d_branca,'step="0.5" min="0"')}</div>
        <div class="form-field"><label class="form-label">L br <span class="unit">(m)</span></label>${makeInput('l_branca','number','L br.',d.l_branca,'step="0.5" min="0"')}</div>
      </div>
      <div class="form-row cols-3">
        <div class="form-field"><label class="form-label">H br <span class="unit">(m)</span></label>${makeInput('h_branca','number','H br.',d.h_branca,'step="0.5" min="0"')}</div>
        <div class="form-field"><label class="form-label">H bers <span class="unit">(m)</span></label>${makeInput('h_bersaglio','number','H bers.',d.h_bersaglio,'step="0.5" min="0"')}</div>
        <div class="form-field"></div>
      </div>
      ${ecoBar}
    </div>
  </div>
  <div class="form-section">
    <div class="form-section-header"><span class="section-icon">⚡</span>Valutazione Bersaglio e Impulso</div>
    <div class="form-grid">
      <div class="form-row cols-2">
        <div class="form-field"><label class="form-label">Tipo Bersaglio</label>${makeSelect('tipo_bersaglio',DATA.tipoBersaglio,'Tipo bersaglio...',d.tipo_bersaglio)}</div>
        <div class="form-field"><label class="form-label">Classe Bersaglio <span class="unit">(1-7)</span></label>${makeInput('classe_bersaglio','number','1-7',d.classe_bersaglio,'min="1" max="7"')}</div>
      </div>
      <div class="impulso-grid">
        <div class="impulso-col">
          <div class="impulso-title">🌲 Bersaglio Albero (B)</div>
          <div class="impulso-row"><span>Impulso (I)</span><span class="mono" id="imp-alb-${ft}">${imp_alb!==null?fmtN(imp_alb,0):'—'}</span><span>kgm/s</span></div>
          <div class="impulso-row"><span>Classe</span><span class="mono" id="cls-alb-${ft}">${cls_alb}</span></div>
        </div>
        <div class="impulso-col">
          <div class="impulso-title">🌿 Bersaglio Rami (B)</div>
          <div class="impulso-row"><span>Impulso (I)</span><span class="mono" id="imp-ram-${ft}">${imp_ram!==null?fmtN(imp_ram,0):'—'}</span><span>kgm/s</span></div>
          <div class="impulso-row"><span>Classe</span><span class="mono" id="cls-ram-${ft}">${cls_ram}</span></div>
        </div>
      </div>
    </div>
  </div>`;
}

// ===== DANGER BTNS =====
// Hints brevi per pericolo ORD (B4:C14 — probabilità cedimento)
const DHINTS={
  '1':"Gravissimi – P cedimento 1/1÷1/5 /anno",
  '2':"Gravi – P cedimento 1/5÷1/50 /anno",
  '3':"Significativi – P cedimento 1/50÷1/500 /anno",
  '4':"Incerti – P cedimento 1/500÷1/5K /anno",
  '5':"Moderati – P cedimento 1/5K÷1/50K /anno",
  '6':"Bassi – P cedimento 1/50K÷1/500K /anno",
  '7':"Trascurabili – P cedimento 1/500K÷1/1M /anno"
};
// Hints per grado pericolo TRIAGE (B17:C23)
const DHINTS_TRG={
  '1':'Molto gravi – connessi con cedimento immediato',
  '2':'Gravi – connessi con cedimento imminente',
  '3':'Significativi – cedimento in cond. critiche',
  '4':'Incerti – occasionalmente connessi con cedimento',
  '5':'Moderati – non chiaramente connessi con cedimento',
  '6':'Modesti – non connessi con cedimento',
  '7':'Trascurabili o sostanzialmente assenti'
};
function buildDangerBtns(field,ft,cur){
  return [1,2,3,4,5,6,7].map(v=>`<button type="button" class="danger-btn ${String(cur)===String(v)?'active':''}" data-val="${v}" onclick="selectDanger(this,${v},'${field}','${ft}')">${v}</button>`).join('');
}
function selectDanger(btn,val,field,ft){
  btn.closest('.danger-selector').querySelectorAll('.danger-btn').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active'); formData[ft][field]=val;
  const hints = (ft==='ord') ? DHINTS : DHINTS_TRG;
  const h=sel('hint-'+field+'-'+ft); if(h) h.textContent=hints[val]||'';
  updateComputed(ft);
}

// ===== CHIPS =====
function buildChips(name,items,selected){
  return items.map(item=>{
    const s=selected.includes(item);
    return `<label class="check-chip ${s?'selected':''}" onclick="toggleChip(this,'${name}','${item}')"><input type="checkbox" ${s?'checked':''}>${item}</label>`;
  }).join('');
}
function toggleChip(el,name,item){
  el.classList.toggle('selected');
  const tab=currentTab;
  if(!formData[tab][name]) formData[tab][name]=[];
  const arr=formData[tab][name],idx=arr.indexOf(item);
  if(idx>=0) arr.splice(idx,1); else arr.push(item);
}

// ===== UPDATE COMPUTED =====
function updateComputed(ft){
  const d=formData[ft];
  const circ=calcCirc(d.d_tronco);
  const bio=calcBio(d.d_tronco,d.h_albero);
  const co2=calcCO2(bio); const o2=calcO2(co2); const ia=calcIA(bio);
  const upd=(id,v,dec=0)=>{const el=sel(id);if(el)el.textContent=v!==null?fmtN(v,dec):'—';};
  upd('circ-'+ft,circ,1);
  upd('bio-'+ft,bio,0); upd('co2-'+ft,co2,0); upd('o2-'+ft,o2,0); upd('ia-'+ft,ia,0);
  const imp_alb=calcImpulsoAlbero(circ,d.h_albero,d.h_bersaglio);
  const imp_ram=calcImpulsoRamo(d.d_branca,d.l_branca,d.h_branca,d.h_bersaglio);
  upd('imp-alb-'+ft,imp_alb,0); upd('imp-ram-'+ft,imp_ram,0);
  const ia_el=sel('cls-alb-'+ft); if(ia_el) ia_el.textContent=classeImpulso(imp_alb);
  const ir_el=sel('cls-ram-'+ft); if(ir_el) ir_el.textContent=classeImpulso(imp_ram);
  const rEl=sel('result-'+ft); if(rEl) rEl.innerHTML=riskHTML(ft);
  if(ft==='trgs'||ft==='ord'){
    const ve=calcValoreEcologico(bio,co2,o2,ia);
    const vc=calcValoreComplessivo(ve,bio,d.condiz_salute);
    const veEl=sel('val-ecol-'+ft); if(veEl) veEl.textContent=fmtEuro(ve);
    const vcEl=sel('val-comp-'+ft); if(vcEl) vcEl.textContent=fmtEuro(vc);
  }
}
function riskHTML(ft){
  const d=formData[ft];
  const p=['branca','tronco','colletto','zolla'].map(k=>parseInt(d[k])).filter(v=>!isNaN(v)&&v>0);
  if(!p.length) return '<span style="color:#999;font-size:13px">Seleziona il grado di pericolo</span>';
  const minP=Math.min(...p), b=parseInt(d.classe_bersaglio);
  if(!b) return '<span style="color:#999;font-size:13px">Inserisci classe bersaglio (1-7)</span>';
  const rc=RISK_TABLE[minP+'-'+b]||Math.ceil((minP+b)/2);
  const info=RISK_LABELS[rc]||RISK_LABELS[4];
  // Per TRG mostra anche l'etichetta triage (EMERGENZA/URGENZA/ecc.)
  const triageLabel = (ft!=='ord') ? `<br><small style="color:${info.color};font-weight:700">${DATA.triageRischio[rc]||''}</small>` : '';
  return `<div class="risk-badge ${info.cls}"><span style="font-size:20px">⚠</span><div><strong>Classe ${rc}</strong> – ${info.label}${triageLabel}<br><small>${info.sub}</small></div></div>`;
}
function onFieldChange(name,value){
  const tab=currentTab;
  if(formData[tab]!==undefined){formData[tab][name]=value; updateComputed(tab);}
}

// ===== DANGER SECTION HELPER =====
function pericSection(ft){
  const d=formData[ft];
  const labels=['Rottura Branca/Rami','Rottura Tronco/Castello','Rottura Colletto','Rib./Sciv. Zolla Radicale'];
  const fields=['branca','tronco','colletto','zolla'];
  const hints = (ft==='ord') ? DHINTS : DHINTS_TRG;
  return labels.map((l,i)=>`<div class="form-field">
    <label class="form-label">${l}</label>
    <div class="danger-selector">${buildDangerBtns(fields[i],ft,d[fields[i]]||'')}</div>
    <div class="danger-hint" id="hint-${fields[i]}-${ft}">${hints[d[fields[i]]]||''}</div>
  </div>`).join('');
}

// ===== FORM BUILDERS =====
function buildTRGP(){
  const d=formData.trgp;
  return `<div class="form-section">
    <div class="form-section-header"><span class="section-icon">📋</span>Dati Generali</div>
    <div class="form-grid">
      <div class="form-row cols-2">
        <div class="form-field"><label class="form-label">Rilevatore</label>${makeInput('rilevatore','text','Nome rilevatore',d.rilevatore)}</div>
        <div class="form-field"><label class="form-label">Data</label>${makeInput('data','date','',d.data||today())}</div>
      </div>
      <div class="form-field"><label class="form-label">Specie</label>${makeSelect('specie',DATA.specie,'Seleziona specie...',d.specie)}</div>
      <div class="form-row cols-2">
        <div class="form-field"><label class="form-label">ID Albero</label>${makeInput('id_albero','text','ID / N°',d.id_albero)}</div>
        <div class="form-field"><label class="form-label">N° Piante</label>${makeInput('n_piante','number','N°',d.n_piante)}</div>
      </div>
      <div class="form-field"><label class="form-label">Ubicazione</label>${makeInput('ubicazione','text','Via / Luogo',d.ubicazione)}</div>
      <div class="form-row cols-2">
        <div class="form-field"><label class="form-label">Dimora</label>${makeSelect('dimora',DATA.dimora,'Dimora...',d.dimora)}</div>
        <div class="form-field"><label class="form-label">Stadio</label>${makeSelect('stadio',[...DATA.stadio,...DATA.stadioPop],'Stadio...',d.stadio)}</div>
      </div>
      <div class="form-row cols-2">
        <div class="form-field"><label class="form-label">Pos. Sociale</label>${makeSelect('pos_sociale',DATA.posSociale,'Posizione...',d.pos_sociale)}</div>
        <div class="form-field"><label class="form-label">Localizzazione</label>${makeSelect('localiz',DATA.localiz,'Localiz...',d.localiz)}</div>
      </div>
      <div class="form-field"><label class="form-label">Vincoli</label>${makeSelect('vincoli',DATA.vincoli,'Vincoli...',d.vincoli)}</div>
    </div>
  </div>
  ${buildMisureSection(d,'trgp',false)}
  <div class="form-section">
    <div class="form-section-header"><span class="section-icon">⚠️</span>Conflitti (max 4)</div>
    <div class="form-grid"><div class="form-row cols-2">
      ${[1,2,3,4].map(i=>`<div class="form-field"><label class="form-label">Conflitto ${i}</label>${makeSelect('conflitto_'+i,DATA.conflitti,'Tipo...',d['conflitto_'+i])}</div>`).join('')}
    </div></div>
  </div>
  <div class="form-section">
    <div class="form-section-header"><span class="section-icon">🔬</span>Grado di Pericolo Percepito (P)</div>
    <div class="form-grid">${pericSection('trgp')}</div>
  </div>
  <div class="form-section">
    <div class="form-section-header"><span class="section-icon">🧪</span>Condizioni di Salute e Vigoria</div>
    <div class="form-grid">
      <div class="form-field"><label class="form-label">Condizioni</label>${makeSelect('condiz_salute',DATA.condizSalute,'Condizioni...',d.condiz_salute)}</div>
    </div>
  </div>
  <div class="form-section">
    <div class="form-section-header"><span class="section-icon">📊</span>Triage del Rischio (R)</div>
    <div class="form-grid"><div class="form-field"><div class="computed-field highlight" id="result-trgp">${riskHTML('trgp')}</div></div></div>
  </div>
  <div class="form-section">
    <div class="form-section-header"><span class="section-icon">✂️</span>Prescrizioni Indicative ed Eventuali Prescrizioni Urgenti</div>
    <div class="form-grid">
      <div class="presc-card-list">
        ${[1,2,3].map(i=>{
          const ki=`presc_int_${i}`,ku=`presc_urg_${i}`,km=`presc_mit_${i}`,kv=`presc_val_${i}`,krm=`presc_urgm_${i}`;
          return `<div class="presc-card">
            <div class="presc-card-num">Prescrizione ${i}</div>
            <div class="presc-card-row">
              <div class="presc-card-label">Intervento Colturale</div>
              <div class="presc-card-value"><select class="form-select" name="${ki}" onchange="onFieldChange('${ki}',this.value)"><option value=""></option>${DATA.prescColturali.map(o=>`<option value="${o}" ${d[ki]===o?'selected':''}>${o}</option>`).join('')}</select></div>
            </div>
            <div class="presc-card-row">
              <div class="presc-card-label">Urgenza</div>
              <div class="presc-card-value"><select class="form-select" name="${ku}" onchange="onFieldChange('${ku}',this.value)"><option value=""></option>${DATA.urgenza.map(o=>`<option value="${o}" ${d[ku]===o?'selected':''}>${o}</option>`).join('')}</select></div>
            </div>
            <div class="presc-card-row">
              <div class="presc-card-label">Mitigazione Bersaglio</div>
              <div class="presc-card-value"><select class="form-select" name="${km}" onchange="onFieldChange('${km}',this.value)"><option value=""></option>${DATA.mitigazBersaglio.map(o=>`<option value="${o}" ${d[km]===o?'selected':''}>${o}</option>`).join('')}</select></div>
            </div>
            <div class="presc-card-row">
              <div class="presc-card-label">Prescr. Valutative</div>
              <div class="presc-card-value"><select class="form-select" name="${kv}" onchange="onFieldChange('${kv}',this.value)"><option value=""></option>${DATA.prescValutative.map(o=>`<option value="${o}" ${d[kv]===o?'selected':''}>${o}</option>`).join('')}</select></div>
            </div>
            <div class="presc-card-row">
              <div class="presc-card-label">Urg/Monit</div>
              <div class="presc-card-value"><select class="form-select" name="${krm}" onchange="onFieldChange('${krm}',this.value)"><option value=""></option>${DATA.monitoraggio.map(o=>`<option value="${o}" ${d[krm]===o?'selected':''}>${o}</option>`).join('')}</select></div>
            </div>
          </div>`;
        }).join('')}
      </div>
    </div>
  </div>
  <div class="form-section">
    <div class="form-section-header"><span class="section-icon">📝</span>Note</div>
    <div class="form-grid"><div class="form-field"><textarea class="form-textarea" name="note" placeholder="Note aggiuntive..." oninput="onFieldChange('note',this.value)">${d.note||''}</textarea></div></div>
  </div>`;
}

function buildTRGS(){
  const d=formData.trgs;
  return `<div class="form-section">
    <div class="form-section-header"><span class="section-icon">📋</span>Dati Generali</div>
    <div class="form-grid">
      <div class="form-row cols-2">
        <div class="form-field"><label class="form-label">Rilevatore</label>${makeInput('rilevatore','text','Nome rilevatore',d.rilevatore)}</div>
        <div class="form-field"><label class="form-label">Data</label>${makeInput('data','date','',d.data||today())}</div>
      </div>
      <div class="form-field"><label class="form-label">Specie</label>${makeSelect('specie',DATA.specie,'Seleziona specie...',d.specie)}</div>
      <div class="form-row cols-2">
        <div class="form-field"><label class="form-label">ID Albero</label>${makeInput('id_albero','text','ID / N°',d.id_albero)}</div>
        <div class="form-field"><label class="form-label">Coer. Fito-Clim</label>${makeSelect('coer_fito',DATA.coerFito,'Coerenza...',d.coer_fito)}</div>
      </div>
      <div class="form-field"><label class="form-label">Ubicazione</label>${makeInput('ubicazione','text','Via / Luogo',d.ubicazione)}</div>
      <div class="form-row cols-2">
        <div class="form-field"><label class="form-label">Dimora</label>${makeSelect('dimora',DATA.dimora,'Dimora...',d.dimora)}</div>
        <div class="form-field"><label class="form-label">Stadio</label>${makeSelect('stadio',[...DATA.stadio,...DATA.stadioPalme],'Stadio...',d.stadio)}</div>
      </div>
      <div class="form-row cols-2">
        <div class="form-field"><label class="form-label">Pos. Sociale</label>${makeSelect('pos_sociale',DATA.posSociale,'Posizione...',d.pos_sociale)}</div>
        <div class="form-field"><label class="form-label">Localizzazione</label>${makeSelect('localiz',DATA.localiz,'Localiz...',d.localiz)}</div>
      </div>
      <div class="form-field"><label class="form-label">Vincoli</label>${makeSelect('vincoli',DATA.vincoli,'Vincoli...',d.vincoli)}</div>
    </div>
  </div>
  ${buildMisureSection(d,'trgs',true)}
  <div class="form-section">
    <div class="form-section-header"><span class="section-icon">⚠️</span>Conflitti (max 2)</div>
    <div class="form-grid"><div class="form-row cols-2">
      ${[1,2].map(i=>`<div class="form-field"><label class="form-label">Conflitto ${i}</label>${makeSelect('conflitto_'+i,DATA.conflitti,'Tipo...',d['conflitto_'+i])}</div>`).join('')}
    </div></div>
  </div>
  <div class="form-section">
    <div class="form-section-header"><span class="section-icon">🔬</span>Grado di Pericolo Percepito (P)</div>
    <div class="form-grid">${pericSection('trgs')}</div>
  </div>
  <div class="form-section">
    <div class="form-section-header"><span class="section-icon">🧪</span>Condizioni di Salute e Vigoria</div>
    <div class="form-grid">
      <div class="form-field"><label class="form-label">Condizioni</label>${makeSelect('condiz_salute',DATA.condizSalute,'Condizioni...',d.condiz_salute)}</div>
    </div>
  </div>
  <div class="form-section">
    <div class="form-section-header"><span class="section-icon">📊</span>Triage del Rischio (R)</div>
    <div class="form-grid"><div class="form-field"><div class="computed-field highlight" id="result-trgs">${riskHTML('trgs')}</div></div></div>
  </div>
  <div class="form-section">
    <div class="form-section-header"><span class="section-icon">✂️</span>Prescrizioni Indicative ed Eventuali Prescrizioni Urgenti</div>
    <div class="form-grid">
      <div class="presc-card-list">
        ${[1,2,3].map(i=>{
          const ki=`presc_int_${i}`,ku=`presc_urg_${i}`,km=`presc_mit_${i}`,kv=`presc_val_${i}`,krm=`presc_urgm_${i}`;
          return `<div class="presc-card">
            <div class="presc-card-num">Prescrizione ${i}</div>
            <div class="presc-card-row">
              <div class="presc-card-label">Intervento Colturale</div>
              <div class="presc-card-value"><select class="form-select" name="${ki}" onchange="onFieldChange('${ki}',this.value)"><option value=""></option>${DATA.prescColturali.map(o=>`<option value="${o}" ${d[ki]===o?'selected':''}>${o}</option>`).join('')}</select></div>
            </div>
            <div class="presc-card-row">
              <div class="presc-card-label">Urgenza</div>
              <div class="presc-card-value"><select class="form-select" name="${ku}" onchange="onFieldChange('${ku}',this.value)"><option value=""></option>${DATA.urgenza.map(o=>`<option value="${o}" ${d[ku]===o?'selected':''}>${o}</option>`).join('')}</select></div>
            </div>
            <div class="presc-card-row">
              <div class="presc-card-label">Mitigazione Bersaglio</div>
              <div class="presc-card-value"><select class="form-select" name="${km}" onchange="onFieldChange('${km}',this.value)"><option value=""></option>${DATA.mitigazBersaglio.map(o=>`<option value="${o}" ${d[km]===o?'selected':''}>${o}</option>`).join('')}</select></div>
            </div>
            <div class="presc-card-row">
              <div class="presc-card-label">Prescr. Valutative</div>
              <div class="presc-card-value"><select class="form-select" name="${kv}" onchange="onFieldChange('${kv}',this.value)"><option value=""></option>${DATA.prescValutative.map(o=>`<option value="${o}" ${d[kv]===o?'selected':''}>${o}</option>`).join('')}</select></div>
            </div>
            <div class="presc-card-row">
              <div class="presc-card-label">Urg/Monit</div>
              <div class="presc-card-value"><select class="form-select" name="${krm}" onchange="onFieldChange('${krm}',this.value)"><option value=""></option>${DATA.monitoraggio.map(o=>`<option value="${o}" ${d[krm]===o?'selected':''}>${o}</option>`).join('')}</select></div>
            </div>
          </div>`;
        }).join('')}
      </div>
    </div>
  </div>
  <div class="form-section">
    <div class="form-section-header"><span class="section-icon">📝</span>Note</div>
    <div class="form-grid"><div class="form-field"><textarea class="form-textarea" name="note" placeholder="Note aggiuntive..." oninput="onFieldChange('note',this.value)">${d.note||''}</textarea></div></div>
  </div>`;
}

function buildORD(){
  const d=formData.ord;

  // Helper: multi-row diagnosis table (like the Excel grid: up to 10 rows per organ section)
  // Each organ section has: aspetto (dropdown from aspettoOrgano), caratteri (dropdown from organ list), difetti (same list)
  function diagRows(prefix, n, lista) {
    let rows = '';
    for(let i=1;i<=n;i++){
      const asp = `diag_${prefix}_asp_${i}`, car = `diag_${prefix}_car_${i}`;
      rows += `<tr>
        <td><select class="form-select diag-sel" name="${asp}" onchange="onFieldChange('${asp}',this.value)">
          <option value=""></option>${DATA.aspettoOrgano.map(o=>`<option value="${o}" ${d[asp]===o?'selected':''}>${o}</option>`).join('')}
        </select></td>
        <td><select class="form-select diag-sel" name="${car}" onchange="onFieldChange('${car}',this.value)">
          <option value=""></option>${lista.map(o=>`<option value="${o}" ${d[car]===o?'selected':''}>${o}</option>`).join('')}
        </select></td>
      </tr>`;
    }
    return rows;
  }

  // Prescrizioni rows
  function prescRows(prefInt, prefVal, prefMit, n) {
    let rows = '';
    for(let i=1;i<=n;i++){
      const ki=`presc_int_${i}`, kv=`presc_val_${i}`, km=`presc_mit_${i}`, ku=`presc_urg_${i}`, krm=`presc_urgm_${i}`;
      rows += `<tr>
        <td><select class="form-select diag-sel" name="${ki}" onchange="onFieldChange('${ki}',this.value)">
          <option value=""></option>${DATA.prescColturali.map(o=>`<option value="${o}" ${d[ki]===o?'selected':''}>${o}</option>`).join('')}
        </select></td>
        <td><select class="form-select diag-sel" name="${ku}" onchange="onFieldChange('${ku}',this.value)">
          <option value=""></option>${DATA.urgenza.map(o=>`<option value="${o}" ${d[ku]===o?'selected':''}>${o}</option>`).join('')}
        </select></td>
        <td><select class="form-select diag-sel" name="${km}" onchange="onFieldChange('${km}',this.value)">
          <option value=""></option>${DATA.mitigazBersaglio.map(o=>`<option value="${o}" ${d[km]===o?'selected':''}>${o}</option>`).join('')}
        </select></td>
        <td><select class="form-select diag-sel" name="${kv}" onchange="onFieldChange('${kv}',this.value)">
          <option value=""></option>${DATA.prescValutative.map(o=>`<option value="${o}" ${d[kv]===o?'selected':''}>${o}</option>`).join('')}
        </select></td>
        <td><select class="form-select diag-sel" name="${krm}" onchange="onFieldChange('${krm}',this.value)">
          <option value=""></option>${DATA.monitoraggio.map(o=>`<option value="${o}" ${d[krm]===o?'selected':''}>${o}</option>`).join('')}
        </select></td>
      </tr>`;
    }
    return rows;
  }

  return `
  <!-- DATI GENERALI -->
  <div class="form-section">
    <div class="form-section-header"><span class="section-icon">📋</span>Dati Generali</div>
    <div class="form-grid">
      <div class="form-row cols-2">
        <div class="form-field"><label class="form-label">Rilevatore</label>${makeInput('rilevatore','text','Nome rilevatore',d.rilevatore)}</div>
        <div class="form-field"><label class="form-label">Data</label>${makeInput('data','date','',d.data||today())}</div>
      </div>
      <div class="form-field"><label class="form-label">Specie</label>${makeSelect('specie',DATA.specie,'Seleziona specie...',d.specie)}</div>
      <div class="form-row cols-2">
        <div class="form-field"><label class="form-label">ID Albero</label>${makeInput('id_albero','text','ID / N°',d.id_albero)}</div>
        <div class="form-field"><label class="form-label">Coer. Fito-Clim/geo</label>${makeSelect('coer_fito',DATA.coerFito,'Coerenza...',d.coer_fito)}</div>
      </div>
      <div class="form-field"><label class="form-label">Ubicazione</label>${makeInput('ubicazione','text','Via / Luogo',d.ubicazione)}</div>
      <div class="form-row cols-2">
        <div class="form-field"><label class="form-label">Dimora</label>${makeSelect('dimora',DATA.dimora,'Dimora...',d.dimora)}</div>
        <div class="form-field"><label class="form-label">Stadio</label>${makeSelect('stadio',DATA.stadio,'Stadio...',d.stadio)}</div>
      </div>
      <div class="form-row cols-2">
        <div class="form-field"><label class="form-label">Pos. Sociale</label>${makeSelect('pos_sociale',DATA.posSociale,'Posizione...',d.pos_sociale)}</div>
        <div class="form-field"><label class="form-label">Localizzazione</label>${makeSelect('localiz',DATA.localiz,'Localiz...',d.localiz)}</div>
      </div>
      <div class="form-row cols-2">
        <div class="form-field"><label class="form-label">Vincoli</label>${makeSelect('vincoli',DATA.vincoli,'Vincoli...',d.vincoli)}</div>
        <div class="form-field"><label class="form-label">Conflitti (max 2)</label>
          <div style="display:flex;gap:6px">
            ${makeSelect('conflitto_1',DATA.conflitti,'Conf. 1',d.conflitto_1)}
            ${makeSelect('conflitto_2',DATA.conflitti,'Conf. 2',d.conflitto_2)}
          </div>
        </div>
      </div>
    </div>
  </div>

  <!-- MISURE + IMPULSO (shared builder) -->
  ${buildMisureSection(d,'ord',true)}

  <!-- DIAGNOSI -->
  <div class="form-section">
    <div class="form-section-header"><span class="section-icon">🔬</span>Diagnosi</div>
    <div class="form-grid">
      <div class="form-field">
        <label class="form-label">Condizioni di Salute e Vigoria</label>
        ${makeSelect('condiz_salute',DATA.condizSalute,'Condizioni...',d.condiz_salute)}
      </div>
      <div class="form-row cols-2">
        <div class="form-field"><label class="form-label">Fitopatia</label>${makeInput('fitopatia','text','Tipo fitopatia',d.fitopatia)}</div>
        <div class="form-field"><label class="form-label">Agente Cariogeno</label>${makeSelect('agente_cariogeno',DATA.agentiCarie,'Agente...',d.agente_cariogeno)}</div>
      </div>
    </div>
  </div>

  <!-- ZOLLA RADICALE -->
  <div class="form-section">
    <div class="form-section-header"><span class="section-icon">🌱</span>Zolla Radicale</div>
    <div class="diag-table-wrap">
      <table class="diag-table">
        <thead><tr><th>Aspetto</th><th>Caratteri / Difetti</th></tr></thead>
        <tbody>${diagRows('zolla',10,DATA.radici)}</tbody>
      </table>
    </div>
  </div>

  <!-- COLLETTO -->
  <div class="form-section">
    <div class="form-section-header"><span class="section-icon">🔵</span>Colletto</div>
    <div class="diag-table-wrap">
      <table class="diag-table">
        <thead><tr><th>Aspetto</th><th>Caratteri / Difetti</th></tr></thead>
        <tbody>${diagRows('colletto',10,DATA.collettoDiag)}</tbody>
      </table>
    </div>
  </div>

  <!-- FUSTO -->
  <div class="form-section">
    <div class="form-section-header"><span class="section-icon">🪵</span>Fusto</div>
    <div class="diag-table-wrap">
      <table class="diag-table">
        <thead><tr><th>Aspetto</th><th>Caratteri / Difetti</th></tr></thead>
        <tbody>${diagRows('fusto',10,DATA.fustoDiag)}</tbody>
      </table>
    </div>
  </div>

  <!-- CASTELLO / BRANCHE E RAMI -->
  <div class="form-section">
    <div class="form-section-header"><span class="section-icon">🌿</span>Castello / Branche e Rami</div>
    <div class="diag-table-wrap">
      <table class="diag-table">
        <thead><tr><th>Aspetto</th><th>Caratteri / Difetti Castello</th></tr></thead>
        <tbody>${diagRows('castello',5,DATA.castelloDiag)}</tbody>
      </table>
      <table class="diag-table" style="margin-top:8px">
        <thead><tr><th>Aspetto</th><th>Caratteri / Difetti Ramificazione</th></tr></thead>
        <tbody>${diagRows('ramif',5,DATA.ramificazDiag)}</tbody>
      </table>
    </div>
  </div>

  <!-- CHIOMA -->
  <div class="form-section">
    <div class="form-section-header"><span class="section-icon">☁️</span>Chioma</div>
    <div class="diag-table-wrap">
      <table class="diag-table">
        <thead><tr><th>Aspetto</th><th>Caratteri / Difetti</th></tr></thead>
        <tbody>${diagRows('chioma',5,DATA.chiomaDiag)}</tbody>
      </table>
    </div>
  </div>

  <!-- GRADO DI PERICOLO -->
  <div class="form-section">
    <div class="form-section-header"><span class="section-icon">⚠️</span>Grado di Pericolo (P) – Probabilità di Cedimento</div>
    <div class="form-grid">${pericSection('ord')}</div>
  </div>

  <!-- RISCHIO -->
  <div class="form-section">
    <div class="form-section-header"><span class="section-icon">📊</span>Calcolo Rischio Attuale</div>
    <div class="form-grid"><div class="form-field"><div class="computed-field highlight" id="result-ord">${riskHTML('ord')}</div></div></div>
  </div>

  <!-- PRESCRIZIONI -->
  <div class="form-section">
    <div class="form-section-header"><span class="section-icon">✂️</span>Prescrizioni Indicative ed Eventuali Prescrizioni Urgenti</div>
    <div class="form-grid">
      <div class="presc-card-list">
        ${[1,2,3].map(i=>{
          const ki=`presc_int_${i}`,ku=`presc_urg_${i}`,km=`presc_mit_${i}`,kv=`presc_val_${i}`,krm=`presc_urgm_${i}`;
          return `<div class="presc-card">
            <div class="presc-card-num">Prescrizione ${i}</div>
            <div class="presc-card-row">
              <div class="presc-card-label">Intervento Colturale</div>
              <div class="presc-card-value"><select class="form-select" name="${ki}" onchange="onFieldChange('${ki}',this.value)"><option value=""></option>${DATA.prescColturali.map(o=>`<option value="${o}" ${d[ki]===o?'selected':''}>${o}</option>`).join('')}</select></div>
            </div>
            <div class="presc-card-row">
              <div class="presc-card-label">Urgenza</div>
              <div class="presc-card-value"><select class="form-select" name="${ku}" onchange="onFieldChange('${ku}',this.value)"><option value=""></option>${DATA.urgenza.map(o=>`<option value="${o}" ${d[ku]===o?'selected':''}>${o}</option>`).join('')}</select></div>
            </div>
            <div class="presc-card-row">
              <div class="presc-card-label">Mitigazione Bersaglio</div>
              <div class="presc-card-value"><select class="form-select" name="${km}" onchange="onFieldChange('${km}',this.value)"><option value=""></option>${DATA.mitigazBersaglio.map(o=>`<option value="${o}" ${d[km]===o?'selected':''}>${o}</option>`).join('')}</select></div>
            </div>
            <div class="presc-card-row">
              <div class="presc-card-label">Prescr. Valutative</div>
              <div class="presc-card-value"><select class="form-select" name="${kv}" onchange="onFieldChange('${kv}',this.value)"><option value=""></option>${DATA.prescValutative.map(o=>`<option value="${o}" ${d[kv]===o?'selected':''}>${o}</option>`).join('')}</select></div>
            </div>
            <div class="presc-card-row">
              <div class="presc-card-label">Urg/Monit</div>
              <div class="presc-card-value"><select class="form-select" name="${krm}" onchange="onFieldChange('${krm}',this.value)"><option value=""></option>${DATA.monitoraggio.map(o=>`<option value="${o}" ${d[krm]===o?'selected':''}>${o}</option>`).join('')}</select></div>
            </div>
          </div>`;
        }).join('')}
      </div>
    </div>
  </div>

  <!-- NOTE -->
  <div class="form-section">
    <div class="form-section-header"><span class="section-icon">📝</span>Note</div>
    <div class="form-grid"><div class="form-field">
      <textarea class="form-textarea" name="note" placeholder="Note aggiuntive..." oninput="onFieldChange('note',this.value)">${d.note||''}</textarea>
    </div></div>
  </div>`;
}

// ===== NAV =====
function switchTab(tab){
  document.querySelectorAll('.screen').forEach(s=>s.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('active'));
  sel('screen-'+tab).classList.add('active'); sel('tab-'+tab).classList.add('active');
  currentTab=tab;
  if(tab==='archive') renderArchive();
  if(tab==='home') renderHomeSaved();
  document.querySelector('.main-content').scrollTop=0;
}
function newForm(type){editingId=null;formData[type]={data:today()};renderForm(type);switchTab(type);}
function renderForm(type){
  const c=sel('form-'+type); if(!c) return;
  if(type==='trgp') c.innerHTML=buildTRGP();
  else if(type==='trgs') c.innerHTML=buildTRGS();
  else if(type==='ord') c.innerHTML=buildORD();
}
function goHome(){switchTab('home');}

// ===== SAVE =====
function saveForm(type){
  const d=formData[type];
  if(!d.specie&&!d.id_albero&&!d.ubicazione){showToast('Inserisci almeno specie, ID o ubicazione','error');return null;}
  const record={id:editingId||uid(),type,savedAt:new Date().toISOString(),data:{...d}};
  const idx=savedForms.findIndex(f=>f.id===record.id);
  if(idx>=0) savedForms[idx]=record; else savedForms.push(record);
  localStorage.setItem('arete_forms',JSON.stringify(savedForms));
  editingId=record.id; showToast('Scheda salvata!','success'); renderHomeSaved();
  return record.id;
}
function saveAndPDF(type){const id=saveForm(type);if(id) setTimeout(()=>exportPDF(id),300);}

// Numero utente registrato fisso
const N_UTENTE = '0368';

// ===== HOME SAVED =====
function renderHomeSaved(){
  const el=sel('home-saved-list');
  if(!savedForms.length){el.innerHTML=`<div class="empty-saved"><div class="emoji">📂</div><div>Nessuna scheda salvata.<br>Crea la tua prima valutazione!</div></div>`;return;}
  const recent=[...savedForms].sort((a,b)=>b.savedAt.localeCompare(a.savedAt)).slice(0,6);
  el.innerHTML=recent.map(f=>{
    const sp=f.data.specie?f.data.specie.split(' - ').pop():''; const id=f.data.id_albero||f.data.ubicazione||'—'; const dt=f.data.data||f.savedAt.split('T')[0];
    return `<div class="saved-item" onclick="openSaved('${f.id}')">
      <span class="saved-badge badge-${f.type}">${f.type.toUpperCase()}</span>
      <div class="saved-info"><div class="title">${sp||id}</div><div class="meta">${id} · ${dt}</div></div>
      <button class="saved-delete" onclick="deleteSaved(event,'${f.id}')">🗑</button>
    </div>`;
  }).join('');
}

// ===== ARCHIVE =====
function filterArchive(type,el){archiveFilter=type;document.querySelectorAll('.filter-chip').forEach(c=>c.classList.remove('active'));el.classList.add('active');renderArchive();}
function renderArchive(){
  const el=sel('archive-list');
  let forms=[...savedForms].sort((a,b)=>b.savedAt.localeCompare(a.savedAt));
  if(archiveFilter!=='all') forms=forms.filter(f=>f.type===archiveFilter);
  if(!forms.length){el.innerHTML=`<div class="archive-empty"><div class="icon">📂</div>Nessuna scheda trovata.</div>`;return;}
  el.innerHTML=forms.map(f=>{
    const sp=f.data.specie||'—'; const id=f.data.id_albero||'—'; const dt=f.data.data||f.savedAt.split('T')[0]; const ubicaz=f.data.ubicazione||'—';
    return `<div class="archive-card">
      <div class="archive-card-header"><span class="saved-badge badge-${f.type}">${f.type.toUpperCase()}</span><span style="font-size:14px;font-weight:600">${sp.split(' - ')[0]}</span></div>
      <div class="archive-card-body">ID: ${id} · Luogo: ${ubicaz} · Data: ${dt}</div>
      <div class="archive-card-footer">
        <button class="btn-sm primary" onclick="openSaved('${f.id}')">✏️ Modifica</button>
        <button class="btn-sm" onclick="exportPDF('${f.id}')">📄 PDF</button>
        <button class="btn-sm danger" onclick="deleteSaved(null,'${f.id}')">🗑 Elimina</button>
      </div>
    </div>`;
  }).join('');
}
function openSaved(id){const f=savedForms.find(x=>x.id===id);if(!f)return;editingId=id;formData[f.type]={...f.data};renderForm(f.type);switchTab(f.type);}
function deleteSaved(evt,id){if(evt){evt.stopPropagation();evt.preventDefault();}if(!confirm('Eliminare questa scheda?'))return;savedForms=savedForms.filter(f=>f.id!==id);localStorage.setItem('arete_forms',JSON.stringify(savedForms));renderArchive();renderHomeSaved();showToast('Scheda eliminata');}

// ===== EXPORT MODAL =====
function showExportModal(id){
  const modal=sel('export-modal'),content=sel('export-modal-content');
  modal.classList.remove('hidden');
  if(id){
    const f=savedForms.find(x=>x.id===id);if(!f)return;
    const sp=(f.data.specie||'').split(' - ')[0]||'';
    content.innerHTML=`<div class="modal-handle"></div><div class="modal-title">Esporta Scheda</div>
      <div class="modal-subtitle">${f.type.toUpperCase()} · ${sp} · ${f.data.data||''}</div>
      <div class="export-options">
        <div class="export-option" id="export-opt-pdf-${id}" onclick="exportPDF('${id}')"><div class="export-icon">📄</div><div class="export-info"><h4>Apri PDF</h4><p>Visualizza e stampa la scheda ARETE</p></div></div>
        <div class="export-option" onclick="exportJSON('${id}')"><div class="export-icon">🗂</div><div class="export-info"><h4>JSON</h4><p>Dati strutturati di questa scheda</p></div></div>
      </div>`;
  } else {
    content.innerHTML=`<div class="modal-handle"></div><div class="modal-title">Esporta Archivio</div>
      <div class="modal-subtitle">Esporta tutte le schede salvate</div>
      <div class="export-options">
        <div class="export-option" onclick="exportAllPDF()"><div class="export-icon">📄</div><div class="export-info"><h4>PDF – Tutte le schede</h4><p>Un unico PDF con tutte le schede</p></div></div>
        <div class="export-option" onclick="exportJSON()"><div class="export-icon">🗂</div><div class="export-info"><h4>JSON</h4><p>Tutti i dati in formato JSON</p></div></div>
        <div class="export-option" onclick="exportCSV()"><div class="export-icon">📊</div><div class="export-info"><h4>CSV</h4><p>Tabella compatibile con Excel</p></div></div>
      </div>`;
  }
}
function closeModal(evt){if(evt.target===sel('export-modal'))sel('export-modal').classList.add('hidden');}

// ===== CONDIVIDI =====
function shareForm(type){
  const id=saveForm(type);
  if(id) setTimeout(()=>shareSaved(id),300);
}

async function shareSaved(id){
  const f=savedForms.find(x=>x.id===id);
  if(!f) return;
  sel('export-modal').classList.add('hidden');

  const tlabels={trgp:'TRG-P Popolamenti',trgs:'TRG-S Singolo Albero',ord:'ORD Valutazione Ordinaria'};
  const d=f.data;
  const specie=(d.specie||'').split(' - ')[0]||'';
  const dataVal=d.data||f.savedAt.split('T')[0];
  const idAlb=d.id_albero||'';
  const titolo=`ARETE – ${tlabels[f.type]||f.type.toUpperCase()}`;
  const nomeFile=`ARETE_${f.type.toUpperCase()}_${(idAlb||specie||'scheda').replace(/[^a-zA-Z0-9]/g,'_').slice(0,30)}_${dataVal}.pdf`;

  showToast('Generazione PDF…');

  // Carica jsPDF se non presente
  if(!window.jspdf){
    await new Promise((res,rej)=>{
      const s=document.createElement('script');
      s.src='https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js';
      s.onload=res; s.onerror=rej;
      document.head.appendChild(s);
    });
  }

  try{
    const {jsPDF}=window.jspdf;
    const doc=new jsPDF({unit:'pt',format:'a4',orientation:'portrait'});
    // A4 in pt: 595.28 x 841.89
    const PW=595.28, PH=841.89;
    const ML=20, MR=20, MT=20;
    const CW=PW-ML-MR; // 555.28pt usabili
    let Y=MT;

    // ── Colori ──────────────────────────────────────────────────────────────
    const VERDE=[26,46,26], VERDE_M=[45,107,45], VERDE_CH=[196,230,196];
    const BLUE=[21,101,192], BLUE_CH=[187,222,251];
    const GRIGIO=[240,244,240], BORDO=[170,170,170];
    const BIANCO=[255,255,255], NERO=[0,0,0];
    const P_COL={1:[198,40,40],2:[198,40,40],3:[216,67,21],4:[249,168,37],5:[174,213,129],6:[102,187,106],7:[56,142,60]};

    // ── Helpers ──────────────────────────────────────────────────────────────
    // Nuova pagina se necessario
    function checkY(needed){
      if(Y+needed>PH-30){ doc.addPage(); Y=MT; }
    }

    // Rettangolo colorato con testo centrato (intestazione sezione)
    function sezione(testo, fillRgb, textRgb=[255,255,255], h=14){
      checkY(h);
      doc.setFillColor(...fillRgb);
      doc.rect(ML,Y,CW,h,'F');
      doc.setDrawColor(...BORDO);
      doc.rect(ML,Y,CW,h,'S');
      doc.setFont('helvetica','bold');
      doc.setFontSize(7.5);
      doc.setTextColor(...textRgb);
      doc.text(testo.toUpperCase(), PW/2, Y+h*0.65, {align:'center'});
      doc.setTextColor(...NERO);
      Y+=h;
    }

    // Disegna bordo griglia
    function bordo(x,y,w,h){ doc.setDrawColor(...BORDO); doc.rect(x,y,w,h,'S'); }

    // Cella label (sfondo grigio) + cella valore
    function cellRow(items, rowH=14){
      // items: [{label,value,lw (larghezza label in pt),vw}] che sommano a CW
      checkY(rowH);
      let x=ML;
      items.forEach(({label,value,lw,vw})=>{
        // label bg
        doc.setFillColor(...GRIGIO);
        doc.rect(x,Y,lw,rowH,'F');
        bordo(x,Y,lw,rowH);
        doc.setFont('helvetica','bold'); doc.setFontSize(6.5); doc.setTextColor(...VERDE_M);
        doc.text(String(label||'').toUpperCase(), x+3, Y+rowH*0.67);
        // value
        bordo(x+lw,Y,vw,rowH);
        doc.setFont('helvetica','normal'); doc.setFontSize(8); doc.setTextColor(...NERO);
        const vstr=String(value||'—');
        const lines=doc.splitTextToSize(vstr, vw-4);
        if(lines.length>1){
          // testo multi-riga: ridistribuisci altezza
          lines.forEach((ln,li)=>{
            doc.text(ln, x+lw+3, Y+9+li*9);
          });
        } else {
          doc.text(vstr, x+lw+3, Y+rowH*0.67);
        }
        x+=lw+vw;
      });
      // calcola altezza effettiva (se multiline)
      let maxLines=1;
      items.forEach(({label,value,lw,vw})=>{
        const ls=doc.splitTextToSize(String(value||'—'),vw-4);
        if(ls.length>maxLines) maxLines=ls.length;
      });
      const actualH=maxLines>1?rowH+(maxLines-1)*9:rowH;
      Y+=actualH;
    }

    // Riga 2 colonne uguali
    function row2(l1,v1,l2,v2,lw=60){
      const hw=CW/2;
      cellRow([{label:l1,value:v1,lw,vw:hw-lw},{label:l2,value:v2,lw,vw:hw-lw}]);
    }

    // Riga full width
    function row1(label,value,lw=70){
      cellRow([{label,value,lw,vw:CW-lw}]);
    }

    // Riga 4 colonne misure
    function row4mis(items){
      checkY(14);
      const cw4=CW/4;
      let x=ML;
      items.forEach(([lbl,val])=>{
        const lw=cw4*0.42, vw=cw4-lw;
        doc.setFillColor(...GRIGIO);
        doc.rect(x,Y,lw,14,'F');
        bordo(x,Y,lw,14);
        doc.setFont('helvetica','bold'); doc.setFontSize(6); doc.setTextColor(...VERDE_M);
        doc.text(String(lbl), x+2, Y+9.5);
        bordo(x+lw,Y,vw,14);
        doc.setFont('courier','bold'); doc.setFontSize(8.5); doc.setTextColor(...NERO);
        doc.text(String(val||'—'), x+lw+3, Y+9.5);
        x+=cw4;
      });
      Y+=14;
    }

    // Cella pericolo (label + numero colorato + hint)
    function rowPericolo(label,val,hint){
      const rowH=14;
      checkY(rowH);
      const lw=130, numW=22, hintW=CW-lw-numW;
      // label
      doc.setFillColor(...GRIGIO);
      doc.rect(ML,Y,lw,rowH,'F');
      bordo(ML,Y,lw,rowH);
      doc.setFont('helvetica','bold'); doc.setFontSize(7); doc.setTextColor(...VERDE_M);
      doc.text(label, ML+3, Y+9.5);
      // numero colorato
      const col=P_COL[parseInt(val)]||[200,200,200];
      doc.setFillColor(...col);
      doc.rect(ML+lw,Y,numW,rowH,'F');
      bordo(ML+lw,Y,numW,rowH);
      if(val){
        const light=parseInt(val)<=2;
        doc.setTextColor(light?255:20,light?255:20,light?255:20);
        doc.setFont('helvetica','bold'); doc.setFontSize(11);
        doc.text(String(val), ML+lw+numW/2, Y+10, {align:'center'});
      }
      // hint
      bordo(ML+lw+numW,Y,hintW,rowH);
      doc.setTextColor(80,80,80); doc.setFont('helvetica','normal'); doc.setFontSize(6.5);
      doc.text(String(hint||''), ML+lw+numW+3, Y+9.5, {maxWidth:hintW-4});
      doc.setTextColor(...NERO);
      Y+=rowH;
    }

    // Testo piccolo
    function note(testo,colore=[80,80,80]){
      checkY(10);
      doc.setFont('helvetica','italic'); doc.setFontSize(6.5); doc.setTextColor(...colore);
      const lines=doc.splitTextToSize(testo,CW);
      doc.text(lines,ML,Y+7);
      Y+=lines.length*8;
      doc.setTextColor(...NERO);
    }

    // ── INTESTAZIONE ─────────────────────────────────────────────────────────
    checkY(36);
    const hw1=CW*0.58, hw2=CW*0.26, hw3=CW*0.16;
    // col 1: titolo
    doc.setDrawColor(...VERDE);
    doc.setLineWidth(1.5);
    doc.rect(ML,Y,hw1,36,'S');
    doc.setLineWidth(0.5);
    doc.setFont('helvetica','italic'); doc.setFontSize(6.5); doc.setTextColor(100,100,100);
    doc.text('Scheda di rilevamento per la Valutazione del Rischio Arboreo',ML+5,Y+10);
    doc.setFont('helvetica','bold'); doc.setFontSize(12); doc.setTextColor(...VERDE);
    doc.text(titolo.replace('ARETE – ',''),ML+5,Y+24);
    doc.setFont('helvetica','normal'); doc.setFontSize(6.5); doc.setTextColor(150,150,150);
    doc.text('ver.2.0 / ver.4.0',ML+5,Y+33);
    // col 2: logo ARETE
    doc.setFillColor(...VERDE);
    doc.rect(ML+hw1,Y,hw2,36,'F');
    doc.setFont('helvetica','bold'); doc.setFontSize(18); doc.setTextColor(245,240,232);
    doc.text('ARETE\u00AE',ML+hw1+hw2/2,Y+20,{align:'center'});
    doc.setFont('helvetica','normal'); doc.setFontSize(6); doc.setTextColor(168,200,168);
    doc.text('PROTOCOLLO',ML+hw1+hw2/2,Y+30,{align:'center'});
    // col 3: utente
    doc.setDrawColor(...VERDE);
    doc.setLineWidth(1.5);
    doc.rect(ML+hw1+hw2,Y,hw3,36,'S');
    doc.setLineWidth(0.5);
    doc.setFont('helvetica','normal'); doc.setFontSize(6.5); doc.setTextColor(80,80,80);
    doc.text('Utente registrato n\u00B0',ML+hw1+hw2+hw3/2,Y+13,{align:'center'});
    doc.setFont('helvetica','bold'); doc.setFontSize(14); doc.setTextColor(...VERDE);
    doc.text(N_UTENTE,ML+hw1+hw2+hw3/2,Y+28,{align:'center'});
    doc.setTextColor(...NERO); doc.setLineWidth(0.5);
    Y+=38;

    // ── DATI GENERALI ────────────────────────────────────────────────────────
    sezione('Dati Generali', VERDE_M);
    row2('Rilevatore',d.rilevatore,'Data',d.data,55);
    row1('Specie',d.specie,55);
    row2('ID Albero',d.id_albero,'Dimora',d.dimora,55);
    row2('Localiz.',d.localiz,'Pos. Sociale',d.pos_sociale,55);
    row2('Stadio',d.stadio,'Vincoli',d.vincoli,55);
    row1('Ubicazione',d.ubicazione,55);
    const confl=[d.conflitto_1,d.conflitto_2,d.conflitto_3,d.conflitto_4].filter(Boolean).join(', ')||'—';
    row2('Conflitti',confl,'Coer. Fito',d.coer_fito,55);

    // ── MISURE + ECOLOGIA ────────────────────────────────────────────────────
    sezione('Misure Dendrologiche', VERDE_M);
    const circ=calcCirc(d.d_tronco);
    const bio=calcBio(d.d_tronco,d.h_albero);
    const co2=calcCO2(bio),o2=calcO2(co2),ia=calcIA(bio);
    row4mis([['H (m)',d.h_albero],['D tr (cm)',d.d_tronco],['Circ \u{1F501}',circ?fmtN(circ,1):'—'],['D ch (m)',d.d_chioma]]);
    row4mis([['D br (cm)',d.d_branca],['L br (m)',d.l_branca],['H br (m)',d.h_branca],['H bers (m)',d.h_bersaglio]]);
    row4mis([['Bio (kg)',bio?fmtN(bio,0):'—'],['CO\u2082 (kg/a)',co2?fmtN(co2,0):'—'],['O\u2082 (kg/a)',o2?fmtN(o2,0):'—'],['I.A. (kg/a)',ia?fmtN(ia,0):'—']]);

    // ── DIAGNOSI ─────────────────────────────────────────────────────────────
    sezione('Diagnosi', VERDE_M);
    row1('Condiz. Salute e Vigoria',d.condiz_salute,100);
    row2('Fitopatia',d.fitopatia,'Agente Cariogeno',d.agente_cariogeno,60);

    if(f.type==='ord'){
      // Tabella morfologica: 4 colonne ZOLLA|COLLETTO|FUSTO+CASTELLO|BRANCHE+CHIOMA
      const getCells=(pfx,n)=>{
        const out=[];
        for(let i=1;i<=n;i++){
          const a=d[`diag_${pfx}_asp_${i}`]||'',c=d[`diag_${pfx}_car_${i}`]||'';
          if(a||c) out.push({asp:a,car:c});
        }
        return out;
      };
      const dZ=getCells('zolla',10),dC=getCells('colletto',10),dF=getCells('fusto',10);
      const dCa=getCells('castello',5),dR=getCells('ramif',5),dCh=getCells('chioma',5);
      const fusCast=[...dF.map(r=>({...r,prefix:'F'})),...dCa.map(r=>({...r,prefix:'C'}))];
      const branChi=[...dR.map(r=>({...r,prefix:'B'})),...dCh.map(r=>({...r,prefix:'CH'}))];
      const nDiag=Math.max(dZ.length,dC.length,fusCast.length,branChi.length);

      if(nDiag>0){
        // intestazioni 4 col
        const cw4=CW/4;
        checkY(12);
        ['ZOLLA RADICALE','COLLETTO','FUSTO & CASTELLO','BRANCHE & CHIOMA'].forEach((t,i)=>{
          doc.setFillColor(220,220,220);
          doc.rect(ML+i*cw4,Y,cw4,12,'F');
          bordo(ML+i*cw4,Y,cw4,12);
          doc.setFont('helvetica','bold'); doc.setFontSize(6.5); doc.setTextColor(...VERDE_M);
          doc.text(t,ML+i*cw4+cw4/2,Y+8,{align:'center'});
        });
        Y+=12;

        // righe diagnosi
        const colPfx={F:[21,101,192],C:[143,58,0],B:[26,107,26],CH:[143,0,110]};
        for(let i=0;i<nDiag;i++){
          const rowH=12;
          checkY(rowH);
          [dZ[i],dC[i],fusCast[i],branChi[i]].forEach((r,ci)=>{
            const x=ML+ci*cw4;
            bordo(x,Y,cw4,rowH);
            if(r){
              const pc=colPfx[r.prefix]||VERDE_M;
              doc.setFont('helvetica','bold'); doc.setFontSize(5.5); doc.setTextColor(...pc);
              if(r.prefix) doc.text(`[${r.prefix}]`,x+2,Y+7.5);
              doc.setFont('helvetica','normal'); doc.setFontSize(6.5); doc.setTextColor(...NERO);
              const lns=doc.splitTextToSize(r.car||'',cw4-18);
              doc.text(lns[0]||'',x+(r.prefix?16:3),Y+7.5);
            }
          });
          Y+=rowH;
        }
      }
    }

    // ── GRADO DI PERICOLO ────────────────────────────────────────────────────
    sezione('Grado di Pericolo (P) – Probabilità di Cedimento', VERDE_M);
    const hintsFn=(v)=> f.type==='ord'? DHINTS[v]||'' : DHINTS_TRG[v]||'';
    rowPericolo('ROTTURA BRANCA/RAMI',d.branca,hintsFn(d.branca));
    rowPericolo('ROTTURA TRONCO/CASTELLO',d.tronco,hintsFn(d.tronco));
    rowPericolo('ROTTURA COLLETTO',d.colletto,hintsFn(d.colletto));
    rowPericolo('RIB/SCIV ZOLLA RADICALE',d.zolla,hintsFn(d.zolla));

    // ── BERSAGLIO E IMPULSO ──────────────────────────────────────────────────
    sezione('Valutazione Bersaglio e Impulso – Pre Intervento', BLUE, BIANCO);
    const impAlb=calcImpulsoAlbero(circ,d.h_albero,d.h_bersaglio);
    const impRam=calcImpulsoRamo(d.d_branca,d.l_branca,d.h_branca,d.h_bersaglio);
    const clsAlb=classeImpulso(impAlb),clsRam=classeImpulso(impRam);
    // riga: bersaglio albero | classe | molt | impulso | bersaglio rami | classe | impulso rami
    checkY(28);
    const bc=CW/7;
    const hdrs=['BERSAGLIO ALBERO (B)','Classe','Molt.','Impulso (I)','BERSAGLIO RAMI (B)','Classe','Impulso (I)'];
    const vals=[d.tipo_bersaglio||'—', d.classe_bersaglio||'—','1',
      impAlb!==null?fmtN(impAlb,0)+' kgm/s\nCl.'+clsAlb:'err',
      d.tipo_bersaglio||'—', d.classe_bersaglio||'—',
      impRam!==null?fmtN(impRam,0)+' kgm/s\nCl.'+clsRam:'err'];
    hdrs.forEach((h,i)=>{
      doc.setFillColor(...BLUE_CH);
      doc.rect(ML+i*bc,Y,bc,12,'F');
      bordo(ML+i*bc,Y,bc,12);
      doc.setFont('helvetica','bold'); doc.setFontSize(6); doc.setTextColor(...BLUE);
      doc.text(h,ML+i*bc+bc/2,Y+8,{align:'center'});
    });
    Y+=12;
    vals.forEach((v,i)=>{
      bordo(ML+i*bc,Y,bc,16);
      const lns=String(v).split('\n');
      doc.setFont('helvetica','bold'); doc.setFontSize(8); doc.setTextColor(...NERO);
      lns.forEach((ln,li)=>{
        doc.text(ln,ML+i*bc+bc/2,Y+9+li*7,{align:'center',maxWidth:bc-2});
      });
    });
    Y+=16;

    // ── RISCHIO (R) ──────────────────────────────────────────────────────────
    const pVals=['branca','tronco','colletto','zolla'].map(k=>parseInt(d[k])).filter(v=>!isNaN(v)&&v>0);
    const minP=pVals.length?Math.min(...pVals):null;
    const bCls=parseInt(d.classe_bersaglio);

    // Riga pericolo pre intervento (compatta)
    checkY(12);
    const qw=CW/4;
    [['BRANCHE/RAMI',d.branca],['TRONCO/CASTELLO',d.tronco],['COLLETTO',d.colletto],['ZOLLA RADICALE',d.zolla]].forEach(([lbl,v],i)=>{
      const col=P_COL[parseInt(v)]||[200,200,200];
      doc.setFillColor(...GRIGIO); doc.rect(ML+i*qw,Y,qw-18,12,'F'); bordo(ML+i*qw,Y,qw-18,12);
      doc.setFont('helvetica','bold'); doc.setFontSize(6); doc.setTextColor(...VERDE_M);
      doc.text(lbl,ML+i*qw+2,Y+8);
      doc.setFillColor(...col); doc.rect(ML+i*qw+qw-18,Y,18,12,'F'); bordo(ML+i*qw+qw-18,Y,18,12);
      if(v){
        const lt=parseInt(v)<=2;
        doc.setTextColor(lt?255:20,lt?255:20,lt?255:20);
        doc.setFont('helvetica','bold'); doc.setFontSize(11);
        doc.text(String(v),ML+i*qw+qw-9,Y+9.5,{align:'center'});
      }
      doc.setTextColor(...NERO);
    });
    Y+=12;

    // Box rischio colorato
    if(minP&&bCls){
      const rc=RISK_TABLE[minP+'-'+bCls]||Math.ceil((minP+bCls)/2);
      const info=RISK_LABELS[rc]||RISK_LABELS[4];
      const rgb=info.color.startsWith('#')?[parseInt(info.color.slice(1,3),16),parseInt(info.color.slice(3,5),16),parseInt(info.color.slice(5,7),16)]:[100,100,100];
      checkY(18);
      doc.setFillColor(...rgb); doc.rect(ML,Y,CW,18,'F');
      doc.setDrawColor(...rgb); doc.setLineWidth(1.5); doc.rect(ML,Y,CW,18,'S'); doc.setLineWidth(0.5);
      doc.setTextColor(255,255,255); doc.setFont('helvetica','bold'); doc.setFontSize(10);
      const label=f.type==='ord'?`RISCHIO ATTUALE – Classe ${rc} – ${info.label.toUpperCase()}`
        :`TRIAGE RISCHIO (R) – ${(DATA.triageRischio[rc]||'').toUpperCase()}`;
      doc.text(label,PW/2,Y+10,{align:'center',maxWidth:CW-10});
      doc.setFont('helvetica','normal'); doc.setFontSize(8);
      doc.text(info.sub,PW/2,Y+16,{align:'center'});
      doc.setTextColor(...NERO); doc.setDrawColor(...BORDO);
      Y+=20;
    }

    // ── PRESCRIZIONI (ORD) ──────────────────────────────────────────────────
    if(f.type==='ord'){
      const hasPres=[1,2,3].some(i=>d[`presc_int_${i}`]||d[`presc_mit_${i}`]||d[`presc_val_${i}`]);
      if(hasPres){
        sezione('Prescrizioni Indicative ed Eventuali Prescrizioni Urgenti', VERDE_M);
        // header 5 col
        checkY(14);
        const pcols=[{t:'INTERVENTI COLTURALI',w:CW*0.32},{t:'URGENZA',w:CW*0.09},{t:'MITIGAZIONE BERSAGLIO',w:CW*0.23},{t:'PRESCRIZIONI VALUTATIVE',w:CW*0.24},{t:'URG/MONIT',w:CW*0.12}];
        let px=ML;
        pcols.forEach(({t,w})=>{
          doc.setFillColor(...GRIGIO); doc.rect(px,Y,w,14,'F'); bordo(px,Y,w,14);
          doc.setFont('helvetica','bold'); doc.setFontSize(6); doc.setTextColor(...VERDE_M);
          doc.text(t,px+w/2,Y+9,{align:'center'});
          px+=w;
        });
        Y+=14;
        for(let i=1;i<=3;i++){
          const ki=d[`presc_int_${i}`]||'',ku=d[`presc_urg_${i}`]||'',km=d[`presc_mit_${i}`]||'';
          const kv=d[`presc_val_${i}`]||'',krm=d[`presc_urgm_${i}`]||'';
          if(!ki&&!km&&!kv) continue;
          // calcola altezza riga
          const maxL=Math.max(
            doc.splitTextToSize(ki,pcols[0].w-4).length,
            doc.splitTextToSize(km,pcols[2].w-4).length,
            doc.splitTextToSize(kv,pcols[3].w-4).length
          );
          const rH=Math.max(14,maxL*9+4);
          checkY(rH);
          px=ML;
          [[ki,pcols[0].w],[ku,pcols[1].w],[km,pcols[2].w],[kv,pcols[3].w],[krm,pcols[4].w]].forEach(([v,w])=>{
            bordo(px,Y,w,rH);
            doc.setFont('helvetica','normal'); doc.setFontSize(7); doc.setTextColor(...NERO);
            const lns=doc.splitTextToSize(v,w-4);
            doc.text(lns,px+2,Y+9,{lineHeightFactor:1.3});
            px+=w;
          });
          Y+=rH;
        }
      }
    } else {
      // TRG-P e TRG-S: stessa tabella prescrizioni completa
      const hasPres=[1,2,3].some(i=>d[`presc_int_${i}`]||d[`presc_mit_${i}`]||d[`presc_val_${i}`]);
      if(hasPres){
        sezione('Prescrizioni Indicative ed Eventuali Prescrizioni Urgenti', VERDE_M);
        checkY(14);
        const pcols=[{t:'INTERVENTI COLTURALI',w:CW*0.32},{t:'URGENZA',w:CW*0.09},{t:'MITIGAZIONE BERSAGLIO',w:CW*0.23},{t:'PRESCRIZIONI VALUTATIVE',w:CW*0.24},{t:'URG/MONIT',w:CW*0.12}];
        let px=ML;
        pcols.forEach(({t,w})=>{
          doc.setFillColor(...GRIGIO);doc.rect(px,Y,w,14,'F');bordo(px,Y,w,14);
          doc.setFont('helvetica','bold');doc.setFontSize(6);doc.setTextColor(...VERDE_M);
          doc.text(t,px+w/2,Y+9,{align:'center'});
          px+=w;
        });
        Y+=14;
        for(let i=1;i<=3;i++){
          const ki=d[`presc_int_${i}`]||'',ku=d[`presc_urg_${i}`]||'',km=d[`presc_mit_${i}`]||'';
          const kv=d[`presc_val_${i}`]||'',krm=d[`presc_urgm_${i}`]||'';
          if(!ki&&!km&&!kv) continue;
          const maxL=Math.max(
            doc.splitTextToSize(ki,pcols[0].w-4).length,
            doc.splitTextToSize(km,pcols[2].w-4).length,
            doc.splitTextToSize(kv,pcols[3].w-4).length
          );
          const rH=Math.max(14,maxL*9+4);
          checkY(rH);px=ML;
          [[ki,pcols[0].w],[ku,pcols[1].w],[km,pcols[2].w],[kv,pcols[3].w],[krm,pcols[4].w]].forEach(([v,w])=>{
            bordo(px,Y,w,rH);
            doc.setFont('helvetica','normal');doc.setFontSize(7);doc.setTextColor(...NERO);
            doc.text(doc.splitTextToSize(v,w-4),px+2,Y+9,{lineHeightFactor:1.3});
            px+=w;
          });
          Y+=rH;
        }
      }
    }

    // ── POST INTERVENTO (ORD) ───────────────────────────────────────────────
    if(f.type==='ord'){
      sezione('Valutazione Bersaglio e Impulso – Post Intervento', BLUE, BIANCO);
      note('Modifiche dimensionali per il calcolo del rischio residuo (inserire solo se modificate a seguito di prescrizioni colturali)',BORDO);
      row4mis([['H',d.h_albero],['Circ',circ?fmtN(circ,1):'—'],['D ch',d.d_chioma],['D br',d.d_branca]]);
      row4mis([['L Br',d.l_branca],['H br',d.h_branca],['H bers',d.h_bersaglio],['','' ]]);
      // rischio residuo = stesso del pre intervento (senza modifiche)
      if(minP&&bCls){
        const rc=RISK_TABLE[minP+'-'+bCls]||Math.ceil((minP+bCls)/2);
        const info=RISK_LABELS[rc]||RISK_LABELS[4];
        const rgb=info.color.startsWith('#')?[parseInt(info.color.slice(1,3),16),parseInt(info.color.slice(3,5),16),parseInt(info.color.slice(5,7),16)]:[100,100,100];
        checkY(14);
        doc.setFillColor(...GRIGIO); doc.rect(ML,Y,100,14,'F'); bordo(ML,Y,100,14);
        doc.setFont('helvetica','bold'); doc.setFontSize(7); doc.setTextColor(...VERDE_M);
        doc.text('PERICOLO POST INT.',ML+3,Y+9.5);
        bordo(ML+100,Y,CW-100,14);
        doc.setFont('helvetica','bold'); doc.setFontSize(8); doc.setTextColor(...rgb);
        doc.text(`RISCHIO RESIDUO (R): Classe ${rc} – ${info.label}`,ML+103,Y+9.5);
        doc.setTextColor(...NERO);
        Y+=14;
        [['BRANCHE/RAMI',d.branca],['TRONCO/CASTELLO',d.tronco],['COLLETTO',d.colletto],['ZOLLA RADICALE',d.zolla]].forEach(([lbl,v],i)=>{
          const col=P_COL[parseInt(v)]||[200,200,200];
          doc.setFillColor(...GRIGIO); doc.rect(ML+i*qw,Y,qw-18,12,'F'); bordo(ML+i*qw,Y,qw-18,12);
          doc.setFont('helvetica','bold'); doc.setFontSize(6); doc.setTextColor(...VERDE_M);
          doc.text(lbl,ML+i*qw+2,Y+8);
          doc.setFillColor(...col); doc.rect(ML+i*qw+qw-18,Y,18,12,'F'); bordo(ML+i*qw+qw-18,Y,18,12);
          if(v){
            const lt=parseInt(v)<=2;
            doc.setTextColor(lt?255:20,lt?255:20,lt?255:20);
            doc.setFont('helvetica','bold'); doc.setFontSize(11);
            doc.text(String(v),ML+i*qw+qw-9,Y+9.5,{align:'center'});
          }
          doc.setTextColor(...NERO);
        });
        Y+=12;
      }
    }

    // ── NOTE ────────────────────────────────────────────────────────────────
    if(d.note){
      sezione('Note', VERDE_M);
      checkY(20);
      doc.setFillColor(255,255,255); doc.rect(ML,Y,CW,20,'F'); bordo(ML,Y,CW,20);
      doc.setFont('helvetica','normal'); doc.setFontSize(8); doc.setTextColor(...NERO);
      const lns=doc.splitTextToSize(d.note,CW-6);
      doc.text(lns,ML+3,Y+10);
      Y+=Math.max(20,lns.length*9+8);
    }

    // ── PIÈ DI PAGINA ───────────────────────────────────────────────────────
    const nPages=doc.internal.getNumberOfPages();
    for(let p=1;p<=nPages;p++){
      doc.setPage(p);
      doc.setFontSize(6.5); doc.setFont('helvetica','normal'); doc.setTextColor(160,160,160);
      doc.text(`ARETE – ${titolo} – Utente n° ${N_UTENTE}  |  Pag. ${p}/${nPages}  |  ${new Date().toLocaleString('it-IT')}`,PW/2,PH-8,{align:'center'});
    }

    // ── CONDIVISIONE ────────────────────────────────────────────────────────
    const pdfBlob=doc.output('blob');
    const pdfFile=new File([pdfBlob],nomeFile,{type:'application/pdf'});

    if(navigator.share && navigator.canShare && navigator.canShare({files:[pdfFile]})){
      try{
        await navigator.share({
          title:titolo,
          text:`Scheda ARETE: ${specie}${idAlb?' · ID '+idAlb:''} · ${dataVal} · Utente n° ${N_UTENTE}`,
          files:[pdfFile]
        });
        showToast('PDF condiviso!','success');
        return;
      } catch(e){
        if(e.name==='AbortError') return;
      }
    }

    // Fallback: scarica
    const url=URL.createObjectURL(pdfBlob);
    const a=document.createElement('a');
    a.href=url; a.download=nomeFile; a.click();
    setTimeout(()=>URL.revokeObjectURL(url),3000);
    showToast('PDF scaricato — allegalo su WhatsApp o Mail','success');

  } catch(err){
    console.error('jsPDF error:',err);
    exportPDF(id);
    showToast('PDF aperto — Stampa → Salva come PDF','');
  }
}


// ===== PDF — layout fedele al modello originale ARETE =====
function dColor(v){const n=parseInt(v);if(!n)return'#fff';const c=['','#c62828','#d84315','#e65100','#f9a825','#aed581','#66bb6a','#388e3c'];return c[n]||'#fff';}
function dTxt(v){const n=parseInt(v);return n&&n<=2?'#fff':'#111';}
function vv(v){return v||'';}
function mv(v,dec=0){return v!==null&&v!==undefined?fmtN(v,dec):'';}

function buildPDFOrd(f){
  const d=f.data;
  const circ=calcCirc(d.d_tronco);
  const bio=calcBio(d.d_tronco,d.h_albero);
  const co2=calcCO2(bio); const o2=calcO2(co2); const ia=calcIA(bio);
  const imp_alb=calcImpulsoAlbero(circ,d.h_albero,d.h_bersaglio);
  const imp_ram=calcImpulsoRamo(d.d_branca,d.l_branca,d.h_branca,d.h_bersaglio);
  const cls_alb=classeImpulso(imp_alb); const cls_ram=classeImpulso(imp_ram);

  const pVals=['branca','tronco','colletto','zolla'].map(k=>parseInt(d[k])).filter(v=>!isNaN(v)&&v>0);
  const minP=pVals.length?Math.min(...pVals):null;
  const bCls=parseInt(d.classe_bersaglio);
  let rColor='#555',rBg='#f5f5f5',rCls='—',rSub='';
  if(minP&&bCls){
    const rc=RISK_TABLE[minP+'-'+bCls]||Math.ceil((minP+bCls)/2);
    const info=RISK_LABELS[rc]||RISK_LABELS[4];
    rColor=info.color; rBg=info.bg;
    rCls=`Classe ${rc}`; rSub=`${info.label} – ${info.sub}`;
  }

  function getDiag(prefix,n){
    const out=[];
    for(let i=1;i<=n;i++){
      const a=d[`diag_${prefix}_asp_${i}`]||'',c=d[`diag_${prefix}_car_${i}`]||'';
      if(a||c) out.push([a,c]);
    }
    return out;
  }
  const dZolla=getDiag('zolla',10), dColl=getDiag('colletto',10), dFusto=getDiag('fusto',10);
  const dCast=getDiag('castello',5), dRamif=getDiag('ramif',5), dChioma=getDiag('chioma',5);
  const dBranRamif=[...dRamif];

  // Merge Fusto+Castello, Branche+Chioma nelle 4 colonne
  const fusCast=[...dFusto.map(r=>['F',r]),...dCast.map(r=>['C',r])];
  const branChi=[...dBranRamif.map(r=>['B',r]),...dChioma.map(r=>['CH',r])];
  const nDiag=Math.max(dZolla.length,dColl.length,fusCast.length,branChi.length,1);

  const prescRows3=[];
  for(let i=1;i<=3;i++){
    prescRows3.push([d[`presc_int_${i}`]||'',d[`presc_urg_${i}`]||'',d[`presc_mit_${i}`]||'',d[`presc_val_${i}`]||'',d[`presc_urgm_${i}`]||'']);
  }

  // Colori cella pericolo con font più grande
  const pericCell=(v,w='5%')=>`<td style="background:${dColor(v)};color:${dTxt(v)};font-family:monospace;font-weight:900;font-size:13pt;text-align:center;width:${w};border:1px solid #aaa">${vv(v)}</td>`;

  // Riga diagnosi morfologica
  const tdDiag='border:1px solid #ccc;padding:2px 4px;vertical-align:top;font-size:8pt;';
  function diagRow(z,co,fc,bc){
    const cell=(item,color)=>item
      ?`<td style="${tdDiag}"><b style="color:${color};font-size:7pt">[${item[0]}]</b> ${item[1][1]}</td>`
      :`<td style="${tdDiag}"></td>`;
    return `<tr>
      ${z?`<td style="${tdDiag}">${z[0]?`<b style="color:#2d6b2d;font-size:7pt">[${z[0]}]</b> `:''}${z[1]}</td>`:`<td style="${tdDiag}"></td>`}
      ${co?`<td style="${tdDiag}">${co[0]?`<b style="color:#2d6b2d;font-size:7pt">[${co[0]}]</b> `:''}${co[1]}</td>`:`<td style="${tdDiag}"></td>`}
      ${fc?cell(fc, fc[0]==='C'?'#8f3a00':'#1a5aa0'):cell(null,'')}
      ${bc?cell(bc, bc[0]==='CH'?'#8f006e':'#1a6b1a'):cell(null,'')}
    </tr>`;
  }

  // La chiave: il wrapper è 270mm ma la stampa scala a A4 (210mm)
  // Fattore di scala = 210/270 ≈ 0.778 → font effettivo 9pt*0.778 ≈ 7pt sul foglio
  // Ma visivamente molto più leggibile nell'interfaccia e con zoom browser

  const css=`
    *{box-sizing:border-box;margin:0;padding:0}
    html{font-size:9pt}
    body{font-family:Arial,Helvetica,sans-serif;font-size:9pt;color:#111;background:#fff}

    /* Screen: mostra a 270mm con font normali */
    .wrap{width:270mm;margin:0 auto;padding:6mm 7mm}

    table{width:100%;border-collapse:collapse}
    td,th{border:1px solid #999;padding:2px 5px;vertical-align:middle;font-size:8.5pt}
    .sh{background:#c8e6c9;font-weight:700;text-align:center;font-size:8pt;
        letter-spacing:0.07em;text-transform:uppercase;padding:3px;border:1.5px solid #2d6b2d}
    .shb{background:#bbdefb;font-weight:700;text-align:center;font-size:8pt;
         letter-spacing:0.07em;text-transform:uppercase;padding:3px;border:1.5px solid #1565c0;color:#0d2660}
    .lbl{background:#eef2ee;font-weight:700;font-size:8pt;white-space:nowrap}
    .mono{font-family:'Courier New',monospace;font-weight:700;font-size:9pt}
    .sec{background:#ddd;font-weight:700;font-size:8pt;text-align:center;
         padding:2px;border:1px solid #aaa;text-transform:uppercase;letter-spacing:0.05em}
    .rbox{padding:4px 8px;font-weight:700;font-size:8.5pt;
          background:${rBg};color:${rColor};border:2px solid ${rColor};border-radius:3px;display:inline-block}
    .pbtn-wrap{position:fixed;top:12px;right:12px;display:flex;flex-direction:column;gap:6px;z-index:99}
    .pbtn{background:#1a2e1a;color:#fff;border:none;
          padding:8px 16px;border-radius:7px;cursor:pointer;font-weight:700;font-size:13px;
          box-shadow:0 3px 12px rgba(0,0,0,.3);width:100%;text-align:center}
    .pbtn-blue{background:#1565c0}
    .foot{margin-top:5px;font-size:7pt;color:#999;border-top:1px solid #ddd;
          padding-top:3px;display:flex;justify-content:space-between}

    @media screen{.wrap{box-shadow:0 2px 24px rgba(0,0,0,.12)}}

    /* Stampa: scala 270mm → 210mm mantenendo font leggibili */
    @media print{
      .pbtn-wrap{display:none}
      html,body{margin:0;padding:0}
      @page{margin:0;size:A4 portrait}
      .wrap{
        width:270mm;
        margin:0;
        padding:5mm 6mm;
        transform-origin:top left;
        transform:scale(0.777);
      }
    }
  `;

  return `<!DOCTYPE html><html lang="it"><head>
  <meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
  <title>ARETE – ORD – ${vv(d.id_albero)}</title>
  <style>${css}</style></head>
  <body>
  <div class="pbtn-wrap">
    <button class="pbtn" onclick="window.print()">🖨 Stampa / Salva PDF</button>
    <button class="pbtn pbtn-blue" onclick="(function(){
      var blob=new Blob([document.documentElement.outerHTML],{type:'text/html;charset=utf-8'});
      var nome=document.title||'ARETE_scheda';
      var file=new File([blob],nome+'.html',{type:'text/html'});
      if(navigator.share&&navigator.canShare&&navigator.canShare({files:[file]})){
        navigator.share({title:nome,files:[file]}).catch(function(e){if(e.name!=='AbortError'){
          var u=URL.createObjectURL(blob);var a=document.createElement('a');
          a.href=u;a.download=nome+'.html';document.body.appendChild(a);a.click();
          setTimeout(function(){document.body.removeChild(a);URL.revokeObjectURL(u);},1000);
        }});
      } else {
        var u=URL.createObjectURL(blob);var a=document.createElement('a');
        a.href=u;a.download=nome+'.html';document.body.appendChild(a);a.click();
        setTimeout(function(){document.body.removeChild(a);URL.revokeObjectURL(u);},1000);
      }
    })()">📤 Condividi</button>
  </div>
  <div class="wrap">
  <table>

  <!-- ① INTESTAZIONE -->
  <tr>
    <td colspan="12" style="border:none;padding:0">
      <table style="table-layout:fixed;width:100%">
        <tr>
          <td style="border:2px solid #1a2e1a;padding:5px 8px;width:58%">
            <div style="font-size:7.5pt;font-style:italic;color:#666">Scheda di rilevamento per la Valutazione del Rischio Arboreo</div>
            <div style="font-size:12pt;font-weight:800;color:#1a2e1a">VALUTAZIONE del RISCHIO ORDINARIA</div>
            <div style="font-size:7pt;color:#999;margin-top:1px">ver.2.0 &nbsp;/&nbsp; ver.4.0</div>
          </td>
          <td style="border:2px solid #1a2e1a;background:#1a2e1a;text-align:center;padding:6px;width:26%">
            <div style="font-size:15pt;font-weight:900;color:#f5f0e8;font-family:Georgia,serif;letter-spacing:-1px">
              AR<span style="color:#a8c8a8">E</span>TE<span style="color:#a8c8a8;font-size:10pt">®</span>
            </div>
            <div style="font-size:6.5pt;color:#a8c8a8;text-transform:uppercase;letter-spacing:0.07em">PROTOCOLLO</div>
          </td>
          <td style="border:2px solid #1a2e1a;text-align:center;padding:5px;width:16%">
            <div style="font-size:7pt;color:#555">Utente registrato n°</div>
            <div style="font-size:14pt;font-weight:900;color:#1a2e1a">${N_UTENTE}</div>
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <!-- ② DATI GENERALI -->
  <tr><td colspan="12" class="sh">DATI GENERALI</td></tr>
  <tr>
    <td class="lbl" style="width:7%">RILEV.</td>
    <td colspan="2" style="width:18%">${vv(d.rilevatore)}</td>
    <td class="lbl" style="width:5%">DATA</td>
    <td style="width:10%">${vv(d.data)}</td>
    <td class="lbl" style="width:5%">SPECIE</td>
    <td colspan="4">${vv(d.specie)}</td>
    <td class="lbl" style="width:3%">ID</td>
    <td style="width:8%">${vv(d.id_albero)}</td>
  </tr>
  <tr>
    <td class="lbl">DIMORA</td><td>${vv(d.dimora)}</td>
    <td class="lbl">LOCALIZ</td><td>${vv(d.localiz)}</td>
    <td class="lbl">P.Soc.</td><td>${vv(d.pos_sociale)}</td>
    <td class="lbl">STADIO</td><td>${vv(d.stadio)}</td>
    <td class="lbl">Vincoli</td><td colspan="3">${vv(d.vincoli)}</td>
  </tr>
  <tr>
    <td class="lbl">UBICAZ</td><td colspan="3">${vv(d.ubicazione)}</td>
    <td class="lbl" colspan="2">CONFLITTI: ${[d.conflitto_1,d.conflitto_2].filter(Boolean).join(', ')||'—'}</td>
    <td class="lbl">COER.FITO</td><td colspan="5">${vv(d.coer_fito)}</td>
  </tr>

  <!-- ③ MISURE + ECOLOGIA -->
  <tr>
    <td class="lbl">H (m)</td><td class="mono">${vv(d.h_albero)}</td>
    <td class="lbl">D tr (cm)</td><td class="mono">${vv(d.d_tronco)}</td>
    <td class="lbl">Circ 🔁</td><td class="mono">${mv(circ,1)}</td>
    <td class="lbl">D ch (m)</td><td class="mono">${vv(d.d_chioma)}</td>
    <td class="lbl">D br (cm)</td><td class="mono">${vv(d.d_branca)}</td>
    <td class="lbl">L br (m)</td><td class="mono">${vv(d.l_branca)}</td>
  </tr>
  <tr>
    <td class="lbl">H br (m)</td><td class="mono">${vv(d.h_branca)}</td>
    <td class="lbl">H bers (m)</td><td class="mono">${vv(d.h_bersaglio)}</td>
    <td class="lbl">Bio (kg)</td><td class="mono">${mv(bio,0)}</td>
    <td class="lbl">CO₂ (kg/a)</td><td class="mono">${mv(co2,0)}</td>
    <td class="lbl">O₂ (kg/a)</td><td class="mono">${mv(o2,0)}</td>
    <td class="lbl">I.A. (kg/a)</td><td class="mono">${mv(ia,0)}</td>
  </tr>
  ${(()=>{const ve=calcValoreEcologico(bio,co2,o2,ia);const vc=calcValoreComplessivo(ve,bio,d.condiz_salute);if(ve===null)return '';return `<tr>
    <td colspan="2" style="background:#e8f5e9;border:1px solid #81c784;padding:4px 6px"><div style="font-size:7pt;font-weight:700;color:#2e7d32;text-transform:uppercase">Valore ecologico</div><div style="font-size:12pt;font-weight:900;color:#1b5e20;font-family:monospace">${fmtEuro(ve)}</div></td>
    <td colspan="2" style="background:#e3f2fd;border:1px solid #90caf9;padding:4px 6px"><div style="font-size:7pt;font-weight:700;color:#1565c0;text-transform:uppercase">Valore complessivo</div><div style="font-size:12pt;font-weight:900;color:#0d47a1;font-family:monospace">${fmtEuro(vc)}</div></td>
    <td colspan="8" style="font-size:7pt;color:#777;padding:4px 6px;vertical-align:middle">Val.ecol.=Bio×0.55+CO₂×1+O₂×5+I.A.×10<br>Val.compl.=Val.ecol.+Val.ornamentale</td>
  </tr>`})()}

  <!-- ④ DIAGNOSI -->
  <tr><td colspan="12" class="sh">DIAGNOSI</td></tr>
  <tr>
    <td colspan="6" style="padding:3px 5px;font-size:8.5pt">
      <b>CONDIZIONI DI SALUTE E VIGORIA:</b> ${vv(d.condiz_salute)}
    </td>
    <td colspan="6" style="padding:3px 5px;font-size:8.5pt">
      <b>fitopatia:</b> ${vv(d.fitopatia)}&nbsp;&nbsp;&nbsp;
      <b>agente cariogeno:</b> ${vv(d.agente_cariogeno)}
    </td>
  </tr>
  <tr>
    <td colspan="12" style="padding:0;border:none">
      <table style="table-layout:fixed;width:100%;border-collapse:collapse">
        <tr>
          <td class="sec" style="width:25%">ZOLLA RADICALE</td>
          <td class="sec" style="width:25%">COLLETTO</td>
          <td class="sec" style="width:25%">FUSTO &amp; CASTELLO</td>
          <td class="sec" style="width:25%">BRANCHE/RAMI &amp; CHIOMA</td>
        </tr>
        ${Array.from({length:nDiag},(_,i)=>diagRow(dZolla[i],dColl[i],fusCast[i],branChi[i])).join('')}
      </table>
    </td>
  </tr>

  <!-- ⑤ GRADO DI PERICOLO -->
  <tr><td colspan="12" class="sh">GRADO DI PERICOLO (P) – PROBABILITÀ DI CEDIMENTO</td></tr>
  <tr>
    <td colspan="3" class="lbl">ROTTURA BRANCA/RAMI</td>
    ${pericCell(d.branca)}
    <td colspan="8" style="font-size:8pt;color:#444">${DHINTS[d.branca]||''}</td>
  </tr>
  <tr>
    <td colspan="3" class="lbl">ROTTURA TRONCO/CASTELLO</td>
    ${pericCell(d.tronco)}
    <td colspan="8" style="font-size:8pt;color:#444">${DHINTS[d.tronco]||''}</td>
  </tr>
  <tr>
    <td colspan="3" class="lbl">ROTTURA COLLETTO</td>
    ${pericCell(d.colletto)}
    <td colspan="8" style="font-size:8pt;color:#444">${DHINTS[d.colletto]||''}</td>
  </tr>
  <tr>
    <td colspan="3" class="lbl">RIB/SCIV ZOLLA RADICALE</td>
    ${pericCell(d.zolla)}
    <td colspan="8" style="font-size:8pt;color:#444">${DHINTS[d.zolla]||''}</td>
  </tr>

  <!-- ⑥ BERSAGLIO E IMPULSO – PRE INTERVENTO -->
  <tr><td colspan="12" class="shb">VALUTAZIONE BERSAGLIO E IMPULSO – PRE INTERVENTO</td></tr>
  <tr>
    <td colspan="3" style="font-size:8.5pt">
      <b>BERSAGLIO ALBERO (B)</b><br>
      <span style="font-size:8pt">${vv(d.tipo_bersaglio)||'—'}</span>
    </td>
    <td class="lbl" style="text-align:center">Classe</td>
    <td class="mono" style="text-align:center;font-size:12pt;font-weight:900">${vv(d.classe_bersaglio)||'—'}</td>
    <td class="lbl" style="text-align:center">MOLT.</td>
    <td class="mono" style="text-align:center">1</td>
    <td class="lbl">IMPULSO (I)</td>
    <td colspan="2" class="mono">${imp_alb!==null?mv(imp_alb,0):'err'} kgm/s &nbsp; Cl. <b>${cls_alb}</b></td>
    <td colspan="2" style="font-size:8.5pt;padding:3px 5px">
      <b>BERSAGLIO RAMI:</b><br>
      <span class="mono">${imp_ram!==null?mv(imp_ram,0):'err'}</span> kgm/s · Cl.<b>${cls_ram}</b>
    </td>
  </tr>
  <tr>
    <td colspan="3" class="lbl">PERICOLO &amp; RISCHIO (R)</td>
    <td colspan="9" style="padding:4px 8px">
      <div class="rbox">${minP&&bCls?`⚠ ${rCls} – ${rSub}`:'Inserire pericolo e classe bersaglio'}</div>
    </td>
  </tr>
  <tr>
    <td colspan="2" class="lbl" style="font-size:8pt">BRANCHE / RAMI</td>${pericCell(d.branca,'4%')}
    <td colspan="2" class="lbl" style="font-size:8pt">TRONCO / CASTELLO</td>${pericCell(d.tronco,'4%')}
    <td colspan="2" class="lbl" style="font-size:8pt">COLLETTO</td>${pericCell(d.colletto,'4%')}
    <td colspan="2" class="lbl" style="font-size:8pt">ZOLLA RADICALE</td>${pericCell(d.zolla,'4%')}
    <td></td>
  </tr>

  <!-- ⑦ PRESCRIZIONI -->
  <tr><td colspan="12" class="sh">PRESCRIZIONI INDICATIVE ED EVENTUALI PRESCRIZIONI URGENTI</td></tr>
  <tr style="background:#eef2ee">
    <td colspan="4" class="lbl" style="text-align:center">INTERVENTI COLTURALI</td>
    <td class="lbl" style="text-align:center">URGENZA</td>
    <td colspan="2" class="lbl" style="text-align:center">MITIGAZIONE BERSAGLIO</td>
    <td class="lbl" style="text-align:center">URG.</td>
    <td colspan="2" class="lbl" style="text-align:center">PRESCRIZIONI VALUTATIVE</td>
    <td colspan="2" class="lbl" style="text-align:center">URG/MONIT</td>
  </tr>
  ${prescRows3.map(([ki,ku,km,kv,krm])=>`<tr>
    <td colspan="4" style="font-size:8pt;padding:2px 4px">${ki}</td>
    <td style="font-size:8pt;text-align:center">${ku}</td>
    <td colspan="2" style="font-size:8pt;padding:2px 4px">${km}</td>
    <td style="font-size:8pt;text-align:center"></td>
    <td colspan="2" style="font-size:8pt;padding:2px 4px">${kv}</td>
    <td colspan="2" style="font-size:8pt;text-align:center">${krm}</td>
  </tr>`).join('')}

  <!-- ⑧ MODIFICHE POST INTERVENTO -->
  <tr>
    <td colspan="12" style="font-size:7pt;font-style:italic;padding:3px 5px;color:#666;border-top:2px solid #aaa">
      Modifiche dimensionali della chioma o ramo per il calcolo del rischio residuo
      <em>(inserire solo se modificate a seguito di prescrizioni colturali)</em>
    </td>
  </tr>
  <tr style="background:#f5f8f5">
    <td class="lbl">H</td><td class="mono">${vv(d.h_albero)}</td>
    <td class="lbl">Circ</td><td class="mono">${mv(circ,1)}</td>
    <td class="lbl">D ch</td><td class="mono">${vv(d.d_chioma)}</td>
    <td class="lbl">D br</td><td class="mono">${vv(d.d_branca)}</td>
    <td class="lbl">L Br</td><td class="mono">${vv(d.l_branca)}</td>
    <td class="lbl">H br</td><td class="mono">${vv(d.h_branca)}</td>
  </tr>
  <tr><td colspan="12" class="shb">VALUTAZIONE BERSAGLIO E IMPULSO – POST INTERVENTO</td></tr>
  <tr>
    <td colspan="3" style="font-size:8.5pt">
      <b>BERSAGLIO ALBERO (B)</b><br>
      <span style="font-size:8pt">${vv(d.tipo_bersaglio)||'traffico 90 km/h'}</span>
    </td>
    <td class="lbl" style="text-align:center">Classe</td>
    <td class="mono" style="text-align:center;font-size:12pt;font-weight:900">${vv(d.classe_bersaglio)||'—'}</td>
    <td class="lbl" style="text-align:center">MOLT.</td>
    <td class="mono" style="text-align:center">1</td>
    <td class="lbl">IMPULSO (I)</td>
    <td colspan="2" class="mono">${imp_alb!==null?mv(imp_alb,0):'err'} kgm/s</td>
    <td colspan="2" style="font-size:8.5pt;padding:3px 5px">
      <b>BERSAGLIO RAMI:</b><br>
      <span class="mono">${imp_ram!==null?mv(imp_ram,0):'err'}</span> kgm/s · Cl.<b>${cls_ram}</b>
    </td>
  </tr>
  <tr>
    <td colspan="3" class="lbl">PERICOLO POST INT.</td>
    <td colspan="9" style="padding:3px 6px;font-size:8.5pt">
      <b>RISCHIO RESIDUO (R):</b>
      <span style="color:${rColor};font-weight:700"> ${minP&&bCls?`${rCls} – ${rSub}`:''}</span>
    </td>
  </tr>
  <tr>
    <td colspan="2" class="lbl" style="font-size:8pt">BRANCHE / RAMI</td>${pericCell(d.branca,'4%')}
    <td colspan="2" class="lbl" style="font-size:8pt">TRONCO / CASTELLO</td>${pericCell(d.tronco,'4%')}
    <td colspan="2" class="lbl" style="font-size:8pt">COLLETTO</td>${pericCell(d.colletto,'4%')}
    <td colspan="2" class="lbl" style="font-size:8pt">ZOLLA RADICALE</td>${pericCell(d.zolla,'4%')}
    <td></td>
  </tr>

  <!-- ⑨ NOTE -->
  <tr><td colspan="12" class="lbl" style="padding:3px 6px;font-size:8.5pt">NOTE</td></tr>
  <tr><td colspan="12" style="height:28px;padding:4px 6px;font-size:8.5pt">${vv(d.note)}</td></tr>

  </table>
  <div class="foot">
    <span>ARETE – Scheda Valutazione Rischio Arboreo – VALUTAZIONE del RISCHIO ORDINARIA – ver.2.0/ver.4.0</span>
    <span>Generato il ${new Date().toLocaleString('it-IT')}</span>
  </div>
  </div></body></html>`;
}

function buildPDF(f){
  if(f.type==='ord') return buildPDFOrd(f);

  const d=f.data, type=f.type;
  const typeLabels={trgp:'TRIAGE - P (popolamenti arborei)',trgs:'TRIAGE - S (singolo albero)'};

  const circ=calcCirc(d.d_tronco);
  const bio=calcBio(d.d_tronco,d.h_albero);
  const co2=calcCO2(bio); const o2=calcO2(co2); const ia=calcIA(bio);
  const imp_alb=calcImpulsoAlbero(circ,d.h_albero,d.h_bersaglio);
  const imp_ram=calcImpulsoRamo(d.d_branca,d.l_branca,d.h_branca,d.h_bersaglio);
  const cls_alb=classeImpulso(imp_alb); const cls_ram=classeImpulso(imp_ram);

  const pVals=['branca','tronco','colletto','zolla'].map(k=>parseInt(d[k])).filter(v=>!isNaN(v)&&v>0);
  const minP=pVals.length?Math.min(...pVals):null;
  const bCls=parseInt(d.classe_bersaglio);
  let rColor='#888',rBg='#f5f5f5',rText='Dati insufficienti per il calcolo del rischio';
  if(minP&&bCls){
    const rc=RISK_TABLE[minP+'-'+bCls]||Math.ceil((minP+bCls)/2);
    const info=RISK_LABELS[rc]||RISK_LABELS[4];
    rColor=info.color; rBg=info.bg;
    rText=`CLASSE ${rc} – ${info.label.toUpperCase()} – ${info.sub}`;
  }
  const conflitti=[d.conflitto_1,d.conflitto_2,d.conflitto_3,d.conflitto_4].filter(Boolean).join(', ')||'—';
  const pericoloFields=[['ROTTURA BRANCA/RAMI',d.branca],['ROTTURA TRONCO/CASTELLO',d.tronco],['ROTTURA COLLETTO',d.colletto],['RIB/SCIV ZOLLA RADICALE',d.zolla]];
  const triageFields=[['BRANCHE / RAMI',d.branca],['TRONCO / CASTELLO',d.tronco],['COLLETTO',d.colletto],['ZOLLA RADICALE',d.zolla]];

  return `<!DOCTYPE html><html lang="it"><head><meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>ARETE – ${typeLabels[type]} – ${d.id_albero||''}</title>
  <style>
    *{box-sizing:border-box;margin:0;padding:0}
    html{font-size:9.5pt}
    body{font-family:Arial,Helvetica,sans-serif;font-size:9.5pt;background:#fff;color:#111}
    .page{width:260mm;margin:0 auto;padding:6mm 7mm}
    table{width:100%;border-collapse:collapse}
    td,th{border:1px solid #999;padding:3px 5px;vertical-align:middle;font-size:9pt}
    .sh{background:#c8e6c9;font-weight:700;text-align:center;font-size:8.5pt;
        letter-spacing:0.07em;text-transform:uppercase;border:1.5px solid #2d6b2d;padding:3px}
    .lbl{background:#eef2ee;font-weight:700;font-size:8.5pt;white-space:nowrap}
    .mono{font-family:'Courier New',monospace;font-weight:700;font-size:9.5pt}
    .pbtn-wrap{position:fixed;top:12px;right:12px;display:flex;flex-direction:column;gap:6px;z-index:99}
    .pbtn{background:#1a2e1a;color:#fff;border:none;
          padding:8px 16px;border-radius:7px;cursor:pointer;font-weight:700;font-size:13px;
          box-shadow:0 3px 12px rgba(0,0,0,.3);width:100%;text-align:center}
    .pbtn-blue{background:#1565c0}
    .foot{margin-top:6px;font-size:7.5pt;color:#aaa;border-top:1px solid #ddd;
          padding-top:3px;display:flex;justify-content:space-between}
    @media screen{.page{box-shadow:0 2px 24px rgba(0,0,0,.12)}}
    @media print{
      .pbtn-wrap{display:none}
      html,body{margin:0;padding:0}
      @page{margin:0;size:A4 portrait}
      .page{
        width:260mm;margin:0;padding:5mm 6mm;
        transform-origin:top left;
        transform:scale(0.808);
      }
    }
  </style></head>
  <body>
  <div class="pbtn-wrap">
    <button class="pbtn" onclick="window.print()">🖨 Stampa / Salva PDF</button>
    <button class="pbtn pbtn-blue" onclick="(function(){
      var blob=new Blob([document.documentElement.outerHTML],{type:'text/html;charset=utf-8'});
      var nome=document.title||'ARETE_scheda';
      var file=new File([blob],nome+'.html',{type:'text/html'});
      if(navigator.share&&navigator.canShare&&navigator.canShare({files:[file]})){
        navigator.share({title:nome,files:[file]}).catch(function(e){if(e.name!=='AbortError'){
          var u=URL.createObjectURL(blob);var a=document.createElement('a');
          a.href=u;a.download=nome+'.html';document.body.appendChild(a);a.click();
          setTimeout(function(){document.body.removeChild(a);URL.revokeObjectURL(u);},1000);
        }});
      } else {
        var u=URL.createObjectURL(blob);var a=document.createElement('a');
        a.href=u;a.download=nome+'.html';document.body.appendChild(a);a.click();
        setTimeout(function(){document.body.removeChild(a);URL.revokeObjectURL(u);},1000);
      }
    })()">📤 Condividi</button>
  </div>
  <div class="page">
  <table>
    <tr>
      <td colspan="4" style="border:2px solid #1a2e1a;padding:8px 10px">
        <div style="font-size:8pt;font-style:italic;color:#555">Scheda di rilevamento per la Valutazione del Rischio Arboreo</div>
        <div style="font-size:13pt;font-weight:800;color:#1a2e1a">${typeLabels[type]}</div>
        <div style="font-size:7pt;color:#888;margin-top:1px">ver.2.0 / ver.4.0</div>
      </td>
      <td style="border:2px solid #1a2e1a;background:#1a2e1a;text-align:center;padding:6px">
        <div style="font-size:16pt;font-weight:900;color:#f5f0e8;font-family:Georgia,serif;letter-spacing:-1px">AR<span style="color:#a8c8a8">E</span>TE<span style="color:#a8c8a8;font-size:10pt">®</span></div>
        <div style="font-size:6pt;color:#a8c8a8;text-transform:uppercase;letter-spacing:0.06em">PROTOCOLLO</div>
      </td>
      <td style="border:2px solid #1a2e1a;text-align:center;padding:5px">
        <div style="font-size:7pt;color:#555">Utente registrato n°</div>
        <div style="font-size:14pt;font-weight:900;color:#1a2e1a">${N_UTENTE}</div>
      </td>
    </tr>
    <tr><td colspan="6" class="sh">DATI GENERALI</td></tr>
    <tr><td class="lbl">RILEVATORE</td><td colspan="2">${vv(d.rilevatore)}</td><td class="lbl">DATA</td><td colspan="2">${vv(d.data)}</td></tr>
    <tr><td class="lbl">SPECIE</td><td colspan="3">${vv(d.specie)}</td><td class="lbl">ID</td><td>${vv(d.id_albero)}</td></tr>
    <tr><td class="lbl">DIMORA</td><td>${vv(d.dimora)}</td><td class="lbl">LOCALIZ</td><td>${vv(d.localiz)}</td><td class="lbl">P.Soc.</td><td>${vv(d.pos_sociale)}</td></tr>
    <tr><td class="lbl">STADIO</td><td>${vv(d.stadio)}</td><td class="lbl">Vincoli</td><td>${vv(d.vincoli)}</td><td class="lbl">COER.Fito</td><td>${vv(d.coer_fito)}</td></tr>
    <tr><td class="lbl">UBICAZIONE</td><td colspan="3">${vv(d.ubicazione)}</td><td class="lbl">CONFLITTI</td><td style="font-size:8pt">${conflitti}</td></tr>
    <tr><td class="lbl">H</td><td class="mono">${vv(d.h_albero)}</td><td class="lbl">D tr</td><td class="mono">${vv(d.d_tronco)}</td><td class="lbl">Circ 🔁</td><td class="mono">${mv(circ,1)}</td></tr>
    <tr><td class="lbl">D ch</td><td class="mono">${vv(d.d_chioma)}</td><td class="lbl">D br</td><td class="mono">${vv(d.d_branca)}</td><td class="lbl">L br</td><td class="mono">${vv(d.l_branca)}</td></tr>
    <tr><td class="lbl">H br</td><td class="mono">${vv(d.h_branca)}</td><td class="lbl">H bers</td><td class="mono">${vv(d.h_bersaglio)}</td><td colspan="2"></td></tr>
    ${bio!==null?`<tr>
      <td class="lbl">Bio (kg)</td><td class="mono">${mv(bio,0)}</td>
      <td class="lbl">CO₂</td><td class="mono">${mv(co2,0)}</td>
      <td class="lbl">O₂</td><td class="mono">${mv(o2,0)}</td>
    </tr>`:''}
    ${(()=>{const ve=calcValoreEcologico(bio,co2,o2,ia);const vc=calcValoreComplessivo(ve,bio,d.condiz_salute);if(ve===null||bio===null)return '';const veStr=fmtEuro(ve);const vcStr=fmtEuro(vc);return '<tr>'
      +'<td style="background:#e8f5e9;border:1px solid #81c784;padding:4px 6px"><div style="font-size:7pt;font-weight:700;color:#2e7d32;text-transform:uppercase">Valore ecologico</div><div style="font-size:11pt;font-weight:900;color:#1b5e20;font-family:monospace">'+veStr+'</div></td>'
      +'<td colspan="2" style="background:#e3f2fd;border:1px solid #90caf9;padding:4px 6px"><div style="font-size:7pt;font-weight:700;color:#1565c0;text-transform:uppercase">Valore complessivo</div><div style="font-size:11pt;font-weight:900;color:#0d47a1;font-family:monospace">'+vcStr+'</div></td>'
      +'<td colspan="3" style="font-size:7pt;color:#777;padding:4px;vertical-align:middle">Val.ecol.=Bio\u00d70.55+CO\u2082+O\u2082\u00d75+I.A.\u00d710<br>Val.compl.=Val.ecol.+Val.ornamentale</td>'
      +'</tr>';})()}    <tr><td class="lbl">Condiz. Salute</td><td colspan="5" style="font-size:8pt">${vv(d.condiz_salute)}</td></tr>
    <tr><td colspan="6" class="sh">GRADO DI PERICOLO PERCEPITO (P)</td></tr>
    ${pericoloFields.map(([l,v])=>`<tr>
      <td class="lbl" style="width:28%">${l}</td>
      <td style="width:8%;background:${dColor(v)};color:${dTxt(v)};font-family:monospace;font-weight:900;font-size:13pt;text-align:center">${vv(v)}</td>
      <td colspan="4" style="font-size:7.5pt;color:#555">${DHINTS[v]||''}</td>
    </tr>`).join('')}
    <tr><td colspan="6" class="sh">VALUTAZIONE BERSAGLIO E IMPULSO</td></tr>
    <tr>
      <td class="lbl">Tipo Bersaglio</td><td>${vv(d.tipo_bersaglio)}</td>
      <td class="lbl">Classe B</td><td class="mono" style="text-align:center;font-size:13pt">${vv(d.classe_bersaglio)||'—'}</td>
      <td class="lbl">Impulso Albero</td><td class="mono">${imp_alb!==null?mv(imp_alb,0):'err'} kgm/s · Cl.${cls_alb}</td>
    </tr>
    <tr><td colspan="4"></td><td class="lbl">Impulso Rami</td><td class="mono">${imp_ram!==null?mv(imp_ram,0):'err'} kgm/s · Cl.${cls_ram}</td></tr>
    <tr><td colspan="6" class="sh">TRIAGE DEL RISCHIO (R)</td></tr>
    ${triageFields.map(([l,v])=>`<tr>
      <td class="lbl">${l}</td>
      <td style="background:${dColor(v)};color:${dTxt(v)};font-family:monospace;font-weight:900;font-size:12pt;text-align:center">${vv(v)}</td>
      <td colspan="4"></td>
    </tr>`).join('')}
    <tr><td colspan="6" style="padding:8px;background:${rBg};color:${rColor};font-weight:700;border:2px solid ${rColor}">${rText}</td></tr>
    ${(()=>{
      const hasPres=[1,2,3].some(i=>d[`presc_int_${i}`]||d[`presc_mit_${i}`]||d[`presc_val_${i}`]);
      if(!hasPres) return '';
      return `<tr><td colspan="6" class="sh">PRESCRIZIONI INDICATIVE ED EVENTUALI PRESCRIZIONI URGENTI</td></tr>
      <tr>
        <td colspan="2" class="lbl" style="text-align:center;font-size:7.5pt">INTERVENTI COLTURALI</td>
        <td class="lbl" style="text-align:center;font-size:7.5pt;width:9%">URGENZA</td>
        <td class="lbl" style="text-align:center;font-size:7.5pt">MITIGAZIONE BERSAGLIO</td>
        <td class="lbl" style="text-align:center;font-size:7.5pt">PRESCR. VALUTATIVE</td>
        <td class="lbl" style="text-align:center;font-size:7.5pt;width:9%">URG/MONIT</td>
      </tr>
      ${[1,2,3].map(i=>{
        const ki=d[`presc_int_${i}`]||'',ku=d[`presc_urg_${i}`]||'',km=d[`presc_mit_${i}`]||'';
        const kv=d[`presc_val_${i}`]||'',krm=d[`presc_urgm_${i}`]||'';
        if(!ki&&!km&&!kv) return '';
        return `<tr>
          <td colspan="2" style="font-size:7.5pt;padding:3px 4px">${ki}</td>
          <td style="font-size:7.5pt;padding:3px 4px;text-align:center">${ku}</td>
          <td style="font-size:7.5pt;padding:3px 4px">${km}</td>
          <td style="font-size:7.5pt;padding:3px 4px">${kv}</td>
          <td style="font-size:7.5pt;padding:3px 4px;text-align:center">${krm}</td>
        </tr>`;
      }).join('')}`;
    })()}
    <tr><td class="lbl">NOTE</td><td colspan="5" style="min-height:36px;padding:8px">${vv(d.note)}</td></tr>
  </table>
  <div class="foot">
    <span>ARETE – ${typeLabels[type]}</span>
    <span>Generato il ${new Date().toLocaleString('it-IT')}</span>
  </div>
  </div></body></html>`;
}

// Mappa id → blob PDF già generato (per condivisione)
const _pdfBlobs = {};

async function exportPDF(id){
  const f=savedForms.find(x=>x.id===id);if(!f)return;
  sel('export-modal').classList.add('hidden');
  const specie=(f.data.specie||'').split(' - ')[0]||'';
  const dataVal=f.data.data||f.savedAt.split('T')[0];
  const idAlb=f.data.id_albero||'';
  const nomeSug='ARETE_'+f.type.toUpperCase()+'_'+(idAlb||specie||'scheda').replace(/[^a-zA-Z0-9]/g,'_').slice(0,30)+'_'+dataVal;
  const nomeInput=prompt('Nome del file PDF:',nomeSug);
  if(nomeInput===null) return;
  const nomePulito=(nomeInput.trim()||nomeSug).replace(/\.pdf$/i,'').replace(/[<>:"/\\|?*]/g,'_');

  // Genera HTML del PDF
  const html=(f.type==='ord'?buildPDFOrd(f):buildPDF(f))
    .replace(/onclick="window\.print\(\)"/g,'onclick="document.title=\''+nomePulito+'\';window.print()"')
    .replace(/<title>[^<]*/,'<title>'+nomePulito);

  // Salva il blob per la condivisione successiva
  const blob=new Blob([html],{type:'text/html;charset=utf-8'});
  _pdfBlobs[id]={blob, nome:nomePulito};

  // Apri/scarica il PDF
  const w=window.open('','_blank');
  if(w){ w.document.write(html); w.document.close(); }
  else { salvaPDFBlob(blob, nomePulito+'.html'); }

  // Ricostruisci il modal con Condividi visibile — delay più lungo per Android
  setTimeout(()=>{
    const sp=(f.data.specie||'').split(' - ')[0]||'';
    const modal=sel('export-modal');
    const content=sel('export-modal-content');
    content.innerHTML=`
      <div class="modal-handle"></div>
      <div class="modal-title">Esporta Scheda</div>
      <div class="modal-subtitle">${f.type.toUpperCase()} · ${sp} · ${f.data.data||''}</div>
      <div class="export-options">
        <div class="export-option" onclick="exportPDF('${id}')">
          <div class="export-icon">📄</div>
          <div class="export-info"><h4>PDF generato ✅</h4><p>Premi per rigenerare con altro nome</p></div>
        </div>
        <div class="export-option" onclick="condividiPDFGenerato('${id}')">
          <div class="export-icon">📤</div>
          <div class="export-info"><h4>Condividi PDF</h4><p>Invia su WhatsApp, Drive, Mail o salva</p></div>
        </div>
        <div class="export-option" onclick="exportJSON('${id}')">
          <div class="export-icon">🗂</div>
          <div class="export-info"><h4>JSON</h4><p>Dati strutturati di questa scheda</p></div>
        </div>
      </div>`;
    modal.classList.remove('hidden');
    showToast('PDF pronto — premi Condividi per inviarlo','success');
  }, 800);
}

// Condividi il PDF già generato con exportPDF
async function condividiPDFGenerato(id){
  const entry=_pdfBlobs[id];
  if(!entry){showToast('Genera prima il PDF','error');return;}
  sel('export-modal').classList.add('hidden');

  const {blob, nome}=entry;
  const nomeFile=nome+'.html';
  const file=new File([blob],nomeFile,{type:'text/html'});

  // Web Share API (Android/iOS — apre menu nativo: Drive, WA, Mail, Files...)
  if(navigator.share&&navigator.canShare&&navigator.canShare({files:[file]})){
    try{
      await navigator.share({
        title:'Scheda ARETE – '+nome,
        text:'Scheda ARETE – '+nome,
        files:[file]
      });
      showToast('✅ PDF condiviso!','success');
      return;
    }catch(e){if(e.name==='AbortError')return;}
  }

  // Fallback Median APK
  if(typeof medianDownloadBlobUrl==='function'){
    try{
      medianDownloadBlobUrl(URL.createObjectURL(blob),'pdf_'+Date.now(),nomeFile);
      showToast('✅ PDF salvato','success');
      return;
    }catch(e){}
  }

  // Fallback download diretto
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a');
  a.href=url;a.download=nomeFile;a.style.display='none';
  document.body.appendChild(a);a.click();
  setTimeout(()=>{document.body.removeChild(a);URL.revokeObjectURL(url);},1000);
  showToast('✅ PDF scaricato','success');
}
function exportAllPDF(){
  if(!savedForms.length){showToast('Nessuna scheda','error');return;}
  const pages=savedForms.map(f=>buildPDF(f).replace(/<!DOCTYPE[\s\S]*?<body[^>]*>/,'').replace(/<\/body[\s\S]*$/,'').replace(/<button[^>]*pbtn[^>]*>[\s\S]*?<\/button>/,''));
  const w=window.open('','_blank');if(!w){showToast('Consenti i popup','error');return;}
  w.document.write(`<!DOCTYPE html><html lang="it"><head><meta charset="UTF-8"><title>ARETE Archivio</title>
  <style>body{font-family:'Segoe UI',Arial,sans-serif;padding:14px;background:#fff}.pbtn{position:fixed;top:14px;right:14px;background:#1a2e1a;color:#f5f0e8;border:none;padding:9px 18px;border-radius:8px;cursor:pointer;font-weight:700}
  .sep{page-break-after:always;border-bottom:3px dashed #ccc;margin:24px 0;padding-bottom:20px}
  @media print{.pbtn{display:none}@page{margin:1cm;size:A4}}</style></head>
  <body><button class="pbtn" onclick="window.print()">🖨 Stampa / Salva PDF</button>
  ${pages.map((p,i)=>`<div class="${i<pages.length-1?'sep':''}">${p}</div>`).join('')}
  </body></html>`);
  w.document.close();
  sel('export-modal').classList.add('hidden');
  showToast(`PDF con ${savedForms.length} schede aperto`,'success');
}

function exportJSON(id){
  const data=id?savedForms.filter(x=>x.id===id):savedForms;
  const blob=new Blob([JSON.stringify(data,null,2)],{type:'application/json'});
  downloadBlob(blob,'arete_export_'+today()+'.json');
  sel('export-modal').classList.add('hidden'); showToast('JSON esportato','success');
}
function exportCSV(){
  if(!savedForms.length){showToast('Nessuna scheda','error');return;}
  const allKeys=new Set(); savedForms.forEach(f=>Object.keys(f.data).forEach(k=>allKeys.add(k)));
  const keys=['type','savedAt',...allKeys];
  const rows=savedForms.map(f=>keys.map(k=>{
    let v=k==='type'?f.type:k==='savedAt'?f.savedAt:(f.data[k]||'');
    if(Array.isArray(v)) v=v.join('|');
    return '"'+String(v).replace(/"/g,'""')+'"';
  }).join(';'));
  const blob=new Blob(['\uFEFF'+[keys.join(';'),...rows].join('\n')],{type:'text/csv;charset=utf-8'});
  downloadBlob(blob,'arete_export_'+today()+'.csv');
  sel('export-modal').classList.add('hidden'); showToast('CSV esportato','success');
}
function downloadBlob(blob,name){
  const url=URL.createObjectURL(blob),a=document.createElement('a');
  a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);
}

// ── generaPDFBlob: genera PDF con jsPDF (condiviso da shareSaved ed exportPDF) ──
async function generaPDFBlob(f,nomeFile){
  if(!window.jspdf){
    try{
      await new Promise((res,rej)=>{
        const s=document.createElement('script');
        s.src='https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js';
        s.onload=res;s.onerror=rej;document.head.appendChild(s);
      });
    }catch(e){return null;}
  }
  try{
    const {jsPDF}=window.jspdf;
    const doc=new jsPDF({unit:'pt',format:'a4',orientation:'portrait'});
    const PW=595.28,PH=841.89,ML=20,MR=20,MT=20,CW=PW-ML-MR;
    let Y=MT;
    const d=f.data;
    const tlabels={trgp:'TRG-P Popolamenti',trgs:'TRG-S Singolo Albero',ord:'ORD Valutazione Ordinaria'};
    const titolo='ARETE \u2013 '+(tlabels[f.type]||f.type.toUpperCase());
    const specie=(d.specie||'').split(' - ')[0]||'';
    const dataVal=d.data||f.savedAt.split('T')[0];
    const idAlb=d.id_albero||'';
    const VERDE=[26,46,26],VERDE_M=[45,107,45],GRIGIO=[240,244,240];
    const BORDO=[170,170,170],BIANCO=[255,255,255],NERO=[0,0,0],BLUE=[21,101,192];
    const P_COL={1:[198,40,40],2:[198,40,40],3:[216,67,21],4:[249,168,37],5:[174,213,129],6:[102,187,106],7:[56,142,60]};
    function checkY(n){if(Y+n>PH-30){doc.addPage();Y=MT;}}
    function bordo(x,y,w,h){doc.setDrawColor(...BORDO);doc.setLineWidth(0.5);doc.rect(x,y,w,h,'S');}
    function sezione(t,col){checkY(13);doc.setFillColor(...col);doc.rect(ML,Y,CW,13,'F');bordo(ML,Y,CW,13);doc.setFont('helvetica','bold');doc.setFontSize(7.5);doc.setTextColor(255,255,255);doc.text(t.toUpperCase(),PW/2,Y+9,{align:'center'});doc.setTextColor(...NERO);Y+=13;}
    function cellRow(items,rH=13){checkY(rH);let x=ML;items.forEach(({lbl,val,lw,vw})=>{doc.setFillColor(...GRIGIO);doc.rect(x,Y,lw,rH,'F');bordo(x,Y,lw,rH);doc.setFont('helvetica','bold');doc.setFontSize(6.5);doc.setTextColor(...VERDE_M);doc.text(String(lbl||'').toUpperCase(),x+2,Y+rH*0.67);bordo(x+lw,Y,vw,rH);doc.setFont('helvetica','normal');doc.setFontSize(8);doc.setTextColor(...NERO);const ln=doc.splitTextToSize(String(val||'\u2014'),vw-4);doc.text(ln[0]||'',x+lw+2,Y+rH*0.67);x+=lw+vw;});Y+=rH;}
    function row2(l1,v1,l2,v2,lw=55){const hw=CW/2;cellRow([{lbl:l1,val:v1,lw,vw:hw-lw},{lbl:l2,val:v2,lw,vw:hw-lw}]);}
    function row1(l,v,lw=70){cellRow([{lbl:l,val:v,lw,vw:CW-lw}]);}
    function row4(items){checkY(13);const cw4=CW/4;let x=ML;items.forEach(([l,v])=>{const lw=cw4*0.42,vw=cw4-lw;doc.setFillColor(...GRIGIO);doc.rect(x,Y,lw,13,'F');bordo(x,Y,lw,13);doc.setFont('helvetica','bold');doc.setFontSize(6);doc.setTextColor(...VERDE_M);doc.text(String(l),x+2,Y+9);bordo(x+lw,Y,vw,13);doc.setFont('courier','bold');doc.setFontSize(8.5);doc.setTextColor(...NERO);doc.text(String(v||'\u2014'),x+lw+2,Y+9);x+=cw4;});Y+=13;}
    // Intestazione
    checkY(34);const hw1=CW*0.58,hw2=CW*0.26,hw3=CW*0.16;
    doc.setDrawColor(...VERDE);doc.setLineWidth(1.5);doc.rect(ML,Y,hw1,34,'S');doc.setLineWidth(0.5);
    doc.setFont('helvetica','bold');doc.setFontSize(10);doc.setTextColor(...VERDE);doc.text(titolo.replace('ARETE \u2013 ',''),ML+4,Y+20);
    doc.setFillColor(...VERDE);doc.rect(ML+hw1,Y,hw2,34,'F');
    doc.setFont('helvetica','bold');doc.setFontSize(15);doc.setTextColor(245,240,232);doc.text('ARETE',ML+hw1+hw2/2,Y+18,{align:'center'});
    doc.setFont('helvetica','normal');doc.setFontSize(6);doc.setTextColor(168,200,168);doc.text('PROTOCOLLO',ML+hw1+hw2/2,Y+27,{align:'center'});
    doc.setDrawColor(...VERDE);doc.setLineWidth(1.5);doc.rect(ML+hw1+hw2,Y,hw3,34,'S');doc.setLineWidth(0.5);
    doc.setFont('helvetica','bold');doc.setFontSize(12);doc.setTextColor(...VERDE);doc.text(N_UTENTE,ML+hw1+hw2+hw3/2,Y+22,{align:'center'});
    doc.setTextColor(...NERO);Y+=36;
    // Dati generali
    sezione('Dati Generali',VERDE_M);
    row2('Rilevatore',d.rilevatore,'Data',d.data,55);
    row1('Specie',d.specie,55);
    row2('ID Albero',d.id_albero,'Dimora',d.dimora,55);
    row2('Localiz.',d.localiz,'Pos. Sociale',d.pos_sociale,55);
    row2('Stadio',d.stadio,'Vincoli',d.vincoli,55);
    row1('Ubicazione',d.ubicazione,55);
    // Misure
    sezione('Misure Dendrologiche',VERDE_M);
    const circ=calcCirc(d.d_tronco);
    const bio=calcBio(d.d_tronco,d.h_albero);
    const co2=calcCO2(bio),o2=calcO2(co2),ia=calcIA(bio);
    row4([['H (m)',d.h_albero],['D tr (cm)',d.d_tronco],['Circ',circ?fmtN(circ,1):'\u2014'],['D ch (m)',d.d_chioma]]);
    row4([['D br (cm)',d.d_branca],['L br (m)',d.l_branca],['H br (m)',d.h_branca],['H bers (m)',d.h_bersaglio]]);
    row4([['Bio (kg)',bio?fmtN(bio,0):'\u2014'],['CO\u2082 (kg/a)',co2?fmtN(co2,0):'\u2014'],['O\u2082 (kg/a)',o2?fmtN(o2,0):'\u2014'],['I.A. (kg/a)',ia?fmtN(ia,0):'\u2014']]);
    // Valore ecologico e complessivo
    if(f.type==='trgs'||f.type==='ord'){
      const ve=calcValoreEcologico(bio,co2,o2,ia);
      const vc=calcValoreComplessivo(ve,bio,d.condiz_salute);
      if(ve!==null){
        checkY(18);const hw=CW/2;
        doc.setFillColor(232,245,233);doc.rect(ML,Y,hw,18,'F');doc.setDrawColor(129,199,132);doc.rect(ML,Y,hw,18,'S');
        doc.setFont('helvetica','bold');doc.setFontSize(7);doc.setTextColor(46,125,50);doc.text('VALORE ECOLOGICO',ML+3,Y+6);
        doc.setFont('helvetica','bold');doc.setFontSize(11);doc.setTextColor(27,94,32);doc.text(fmtEuro(ve),ML+3,Y+15);
        doc.setFillColor(227,242,253);doc.rect(ML+hw,Y,hw,18,'F');doc.setDrawColor(144,202,249);doc.rect(ML+hw,Y,hw,18,'S');
        doc.setFont('helvetica','bold');doc.setFontSize(7);doc.setTextColor(21,101,192);doc.text('VALORE COMPLESSIVO',ML+hw+3,Y+6);
        doc.setFont('helvetica','bold');doc.setFontSize(11);doc.setTextColor(13,71,161);doc.text(fmtEuro(vc),ML+hw+3,Y+15);
        doc.setTextColor(...NERO);doc.setDrawColor(...BORDO);Y+=20;
      }
    }
    // Diagnosi
    sezione('Diagnosi',VERDE_M);
    row1('Condiz. Salute',d.condiz_salute,80);
    row2('Fitopatia',d.fitopatia,'Agente Cariogeno',d.agente_cariogeno,60);
    // Grado pericolo
    sezione('Grado di Pericolo (P)',VERDE_M);
    [['BRANCA/RAMI',d.branca],['TRONCO/CASTELLO',d.tronco],['COLLETTO',d.colletto],['ZOLLA RADICALE',d.zolla]].forEach(([l,v])=>{
      checkY(13);const lw=120,nw=20,hw2=CW-lw-nw;
      doc.setFillColor(...GRIGIO);doc.rect(ML,Y,lw,13,'F');bordo(ML,Y,lw,13);
      doc.setFont('helvetica','bold');doc.setFontSize(7);doc.setTextColor(...VERDE_M);doc.text(l,ML+2,Y+9);
      const col=P_COL[parseInt(v)]||[200,200,200];
      doc.setFillColor(...col);doc.rect(ML+lw,Y,nw,13,'F');bordo(ML+lw,Y,nw,13);
      if(v){const lt=parseInt(v)<=2;doc.setTextColor(lt?255:20,lt?255:20,lt?255:20);doc.setFont('helvetica','bold');doc.setFontSize(11);doc.text(String(v),ML+lw+nw/2,Y+10,{align:'center'});}
      bordo(ML+lw+nw,Y,hw2,13);
      doc.setTextColor(80,80,80);doc.setFont('helvetica','normal');doc.setFontSize(6.5);
      const hint=f.type==='ord'?(DATA.gradoPericoloOrd||{})[parseInt(v)]||'':(DATA.gradoPericoloTriage||{})[parseInt(v)]||'';
      doc.text(String(hint),ML+lw+nw+2,Y+9,{maxWidth:hw2-4});
      doc.setTextColor(...NERO);Y+=13;
    });
    // Rischio
    const pVals=['branca','tronco','colletto','zolla'].map(k=>parseInt(d[k])).filter(v=>!isNaN(v)&&v>0);
    const minP=pVals.length?Math.min(...pVals):null;
    const bCls=parseInt(d.classe_bersaglio);
    if(minP&&bCls){
      const rc=RISK_TABLE[minP+'-'+bCls]||Math.ceil((minP+bCls)/2);
      const info=RISK_LABELS[rc]||RISK_LABELS[4];
      const rgb=info.color.startsWith('#')?[parseInt(info.color.slice(1,3),16),parseInt(info.color.slice(3,5),16),parseInt(info.color.slice(5,7),16)]:[100,100,100];
      checkY(16);doc.setFillColor(...rgb);doc.rect(ML,Y,CW,16,'F');doc.setDrawColor(...rgb);doc.rect(ML,Y,CW,16,'S');
      doc.setTextColor(255,255,255);doc.setFont('helvetica','bold');doc.setFontSize(9);
      const lbl=f.type==='ord'?('RISCHIO ATTUALE \u2013 Classe '+rc+' \u2013 '+info.label.toUpperCase()):('TRIAGE RISCHIO \u2013 '+(DATA.triageRischio[rc]||'').toUpperCase());
      doc.text(lbl,PW/2,Y+10,{align:'center'});doc.setTextColor(...NERO);doc.setDrawColor(...BORDO);Y+=18;
    }

    // ── VALUTAZIONE BERSAGLIO E IMPULSO ────────────────────────────────────
    sezione('Valutazione Bersaglio e Impulso',VERDE_M);
    const imp_alb=calcImpulsoAlbero(circ,d.h_albero,d.h_bersaglio);
    const imp_ram=calcImpulsoRamo(d.d_branca,d.l_branca,d.h_branca,d.h_bersaglio);
    const cls_alb=classeImpulso(imp_alb);
    const cls_ram=classeImpulso(imp_ram);
    row2('Tipo Bersaglio',d.tipo_bersaglio,'Classe Bersaglio',d.classe_bersaglio,80);
    row4([['Imp.Albero (kgm/s)',imp_alb!==null?fmtN(imp_alb,0):'\u2014'],['Classe',String(cls_alb)],['Imp.Rami (kgm/s)',imp_ram!==null?fmtN(imp_ram,0):'\u2014'],['Classe',String(cls_ram)]]);

    // ── PRESCRIZIONI ────────────────────────────────────────────────────────
    const hasPres=[1,2,3].some(i=>d[`presc_int_${i}`]||d[`presc_mit_${i}`]||d[`presc_val_${i}`]);
    if(hasPres){
      sezione('Prescrizioni Indicative ed Eventuali Prescrizioni Urgenti',VERDE_M);
      checkY(13);
      const pcols=[{t:'INTERVENTI COLTURALI',w:CW*0.34},{t:'URGENZA',w:CW*0.10},{t:'MITIGAZIONE BERSAGLIO',w:CW*0.22},{t:'PRESCR. VALUTATIVE',w:CW*0.22},{t:'URG/MONIT',w:CW*0.12}];
      let px=ML;
      pcols.forEach(({t,w})=>{
        doc.setFillColor(...GRIGIO);doc.rect(px,Y,w,13,'F');bordo(px,Y,w,13);
        doc.setFont('helvetica','bold');doc.setFontSize(5.5);doc.setTextColor(...VERDE_M);
        doc.text(t,px+w/2,Y+8.5,{align:'center'});px+=w;
      });
      Y+=13;
      for(let i=1;i<=3;i++){
        const ki=d[`presc_int_${i}`]||'',ku=d[`presc_urg_${i}`]||'',km=d[`presc_mit_${i}`]||'';
        const kv=d[`presc_val_${i}`]||'',krm=d[`presc_urgm_${i}`]||'';
        if(!ki&&!km&&!kv) continue;
        const maxL=Math.max(
          doc.splitTextToSize(ki,pcols[0].w-3).length,
          doc.splitTextToSize(km,pcols[2].w-3).length,
          doc.splitTextToSize(kv,pcols[3].w-3).length
        );
        const rH=Math.max(13,maxL*8+4);
        checkY(rH);px=ML;
        [[ki,pcols[0].w],[ku,pcols[1].w],[km,pcols[2].w],[kv,pcols[3].w],[krm,pcols[4].w]].forEach(([v,w])=>{
          bordo(px,Y,w,rH);
          doc.setFont('helvetica','normal');doc.setFontSize(6.5);doc.setTextColor(...NERO);
          doc.text(doc.splitTextToSize(v,w-3),px+2,Y+8,{lineHeightFactor:1.25});
          px+=w;
        });
        Y+=rH;
      }
    }

    // Note
    if(d.note){
      sezione('Note',VERDE_M);checkY(18);
      doc.setFillColor(255,255,255);doc.rect(ML,Y,CW,18,'F');bordo(ML,Y,CW,18);
      doc.setFont('helvetica','normal');doc.setFontSize(8);
      doc.text(doc.splitTextToSize(d.note,CW-6),ML+3,Y+9);Y+=18;
    }
    // Piè di pagina
    const nP=doc.internal.getNumberOfPages();
    for(let p=1;p<=nP;p++){
      doc.setPage(p);doc.setFontSize(6.5);doc.setFont('helvetica','normal');doc.setTextColor(160,160,160);
      doc.text('ARETE \u2013 '+titolo+' \u2013 Utente n\u00b0 '+N_UTENTE+'  |  Pag. '+p+'/'+nP+'  |  '+new Date().toLocaleString('it-IT'),PW/2,PH-8,{align:'center'});
    }
    return doc.output('blob');
  }catch(e){console.error('jsPDF:',e);return null;}
}

function salvaPDFBlob(pdfBlob,nomeFile){
  if(typeof medianDownloadBlobUrl==='function'){
    try{medianDownloadBlobUrl(URL.createObjectURL(pdfBlob),'pdf_'+Date.now(),nomeFile);showToast('PDF salvato in Download','success');return;}catch(e){}
  }
  const reader=new FileReader();
  reader.onload=function(e){
    const a=document.createElement('a');a.href=e.target.result;a.download=nomeFile;a.style.display='none';
    document.body.appendChild(a);a.click();setTimeout(()=>document.body.removeChild(a),1000);
    showToast('PDF scaricato','success');
  };
  reader.readAsDataURL(pdfBlob);
}

function downloadDiretto(url,nome){
  const a=document.createElement('a');a.href=url;a.download=nome;a.style.display='none';
  document.body.appendChild(a);a.click();setTimeout(()=>document.body.removeChild(a),1000);
}

function shareSaved(id){
  const f=savedForms.find(x=>x.id===id);
  if(!f) return;
  sel('export-modal').classList.add('hidden');
  const tlabels={trgp:'TRG-P Popolamenti',trgs:'TRG-S Singolo Albero',ord:'ORD Valutazione Ordinaria'};
  const d=f.data;
  const specie=(d.specie||'').split(' - ')[0]||'';
  const dataVal=d.data||f.savedAt.split('T')[0];
  const idAlb=d.id_albero||'';
  const titolo='ARETE \u2013 '+(tlabels[f.type]||f.type.toUpperCase());
  const nomeSuggerito='ARETE_'+f.type.toUpperCase()+'_'+(idAlb||specie||'scheda').replace(/[^a-zA-Z0-9]/g,'_').slice(0,30)+'_'+dataVal;
  mostraMenuCondivisione(f, nomeSuggerito, titolo, specie, dataVal, idAlb);
}

function mostraMenuCondivisione(f, nomeSuggerito, titolo, specie, dataVal, idAlb){
  const old=document.getElementById('share-dialog');
  if(old) old.remove();

  const testo='Scheda ARETE: '+specie+(idAlb?' \u00b7 ID '+idAlb:'')+' \u00b7 '+dataVal+' \u00b7 Utente n\u00b0 '+N_UTENTE;
  const mailSubj=encodeURIComponent(titolo+' \u2013 '+specie+' \u2013 '+dataVal);
  const mailBody=encodeURIComponent(testo+'\n\nIn allegato la scheda in formato PDF.');

  const overlay=document.createElement('div');
  overlay.id='share-dialog';
  overlay.style.cssText='position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,.6);z-index:9000;display:flex;align-items:flex-end;justify-content:center';
  overlay.addEventListener('click',function(e){if(e.target===overlay)overlay.remove();});

  const sheet=document.createElement('div');
  sheet.style.cssText='background:#f5f0e8;border-radius:20px 20px 0 0;padding:20px 16px 32px;width:100%;max-width:480px';

  // Handle
  const handle=document.createElement('div');
  handle.style.cssText='width:36px;height:4px;background:#c8bfaa;border-radius:2px;margin:0 auto 16px';
  sheet.appendChild(handle);

  // Titolo
  const title=document.createElement('div');
  title.style.cssText='font-size:16px;font-weight:700;color:#1a2e1a;margin-bottom:12px';
  title.textContent='Condividi PDF';
  sheet.appendChild(title);

  // Label nome
  const labelNome=document.createElement('div');
  labelNome.style.cssText='font-size:12px;font-weight:600;color:#555;margin-bottom:4px';
  labelNome.textContent='Nome file';
  sheet.appendChild(labelNome);

  // Input nome
  const inputNome=document.createElement('input');
  inputNome.type='text';
  inputNome.value=nomeSuggerito;
  inputNome.style.cssText='width:100%;padding:9px 12px;border:1.5px solid #c8bfaa;border-radius:8px;font-size:13px;background:#fff;box-sizing:border-box;margin-bottom:16px;font-family:inherit';
  sheet.appendChild(inputNome);

  // Opzioni
  const opzioni=[
    {act:'drive',    label:'Google Drive',  sub:'Carica nel tuo Drive'},
    {act:'whatsapp', label:'WhatsApp',       sub:'Invia via WhatsApp'},
    {act:'mail',     label:'Mail',           sub:'Invia per email'},
    {act:'download', label:'Scarica PDF',    sub:'Salva sul dispositivo'},
  ];

  const icone={
    drive:'<svg width="28" height="24" viewBox="0 0 87.3 78"><path d="m6.6 66.85 3.85 6.65c.8 1.4 1.95 2.5 3.3 3.3l13.75-23.8h-27.5c0 1.55.4 3.1 1.2 4.5z" fill="#0066da"/><path d="m43.65 25-13.75-23.8c-1.35.8-2.5 1.9-3.3 3.3l-25.4 44a9.06 9.06 0 0 0-1.2 4.5h27.5z" fill="#00ac47"/><path d="m73.55 76.8c1.35-.8 2.5-1.9 3.3-3.3l1.6-2.75 7.65-13.25c.8-1.4 1.2-2.95 1.2-4.5h-27.5l5.85 11.5z" fill="#ea4335"/><path d="m43.65 25 13.75-23.8c-1.35-.8-2.9-1.2-4.5-1.2h-18.5c-1.6 0-3.15.45-4.5 1.2z" fill="#00832d"/><path d="m59.8 53h-32.3l-13.75 23.8c1.35.8 2.9 1.2 4.5 1.2h50.8c1.6 0 3.15-.45 4.5-1.2z" fill="#2684fc"/><path d="m73.4 26.5-12.7-22c-.8-1.4-1.95-2.5-3.3-3.3l-13.75 23.8 16.15 27h27.45c0-1.55-.4-3.1-1.2-4.5z" fill="#ffba00"/></svg>',
    whatsapp:'<svg width="28" height="28" viewBox="0 0 24 24" fill="#25d366"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347z"/><path d="M12 0C5.373 0 0 5.373 0 12c0 2.115.554 4.1 1.523 5.824L.057 23.8a.5.5 0 00.614.658l6.142-1.612A11.94 11.94 0 0012 24c6.627 0 12-5.373 12-12S18.627 0 12 0zm0 21.818a9.793 9.793 0 01-5.001-1.371l-.357-.213-3.705.972.989-3.614-.233-.372A9.784 9.784 0 012.182 12C2.182 6.57 6.57 2.182 12 2.182S21.818 6.57 21.818 12 17.43 21.818 12 21.818z"/></svg>',
    mail:'<span style="font-size:26px">&#x1F4E7;</span>',
    download:'<span style="font-size:26px">&#x2B07;&#xFE0F;</span>',
  };

  opzioni.forEach(function(o){
    const btn=document.createElement('button');
    btn.type='button';
    btn.style.cssText='display:flex;align-items:center;gap:14px;padding:14px;border:1.5px solid #ddd;border-radius:12px;background:#fff;width:100%;text-align:left;cursor:pointer;margin-bottom:8px;font-family:inherit';
    btn.innerHTML='<span style="width:36px;text-align:center;flex-shrink:0">'+icone[o.act]+'</span>'
      +'<span><strong style="display:block;font-size:14px">'+o.label+'</strong>'
      +'<span style="font-size:12px;color:#888">'+o.sub+'</span></span>';

    btn.addEventListener('click', async function(){
      const nomeFile=(inputNome.value.trim()||nomeSuggerito).replace(/\.pdf$/i,'').replace(/[<>:"\/\\|?*]/g,'_')+'.pdf';
      overlay.remove();
      showToast('Generazione PDF...');
      // Usa buildPDF HTML (stesso del tasto PDF) per avere PDF identici
      const html=buildPDF(f);
      const pdfBlob=new Blob([html],{type:'text/html;charset=utf-8'});
      const pdfFile=new File([pdfBlob],nomeFile.replace(/\.pdf$/,'.html'),{type:'text/html'});
      const pdfUrl=URL.createObjectURL(pdfBlob);
      setTimeout(function(){URL.revokeObjectURL(pdfUrl);},300000);
      if(o.act==='drive'){
        await _doSalvaSuDrive(pdfUrl,nomeFile);
      } else if(o.act==='whatsapp'){
        await _doShare(pdfFile,pdfUrl,nomeFile,testo,'https://wa.me/?text='+encodeURIComponent(testo+'\n\n(allegare il PDF scaricato)'));
      } else if(o.act==='mail'){
        await _doShare(pdfFile,pdfUrl,nomeFile,testo,'mailto:?subject='+mailSubj+'&body='+mailBody);
      } else {
        salvaPDFBlob(pdfBlob,nomeFile);
        showToast('PDF scaricato','success');
      }
    });
    sheet.appendChild(btn);
  });

  overlay.appendChild(sheet);
  document.body.appendChild(overlay);
}

async function _doShare(pdfFile,pdfUrl,nomeFile,testo,fallbackUrl){
  if(navigator.share&&navigator.canShare&&navigator.canShare({files:[pdfFile]})){
    try{
      await navigator.share({title:'ARETE PDF',text:testo,files:[pdfFile]});
      showToast('PDF condiviso!','success');
      return;
    }catch(e){if(e.name==='AbortError')return;}
  }
  salvaPDFBlob(pdfFile,nomeFile);
  try{window.open(fallbackUrl,'_blank');}catch(e){}
  showToast('PDF scaricato - allegalo manualmente','success');
}


// ── Drive upload ────────────────────────────────────────────────────────────
let _driveToken=null,_driveTokenExpiry=0;
async function getDriveToken(){
  if(_driveToken&&Date.now()<_driveTokenExpiry-60000) return _driveToken;
  let cid=localStorage.getItem('arete_gcid')||'';
  if(!cid){cid=prompt('Client ID Google OAuth2:','');if(!cid)return null;localStorage.setItem('arete_gcid',cid.trim());}
  return new Promise(resolve=>{
    const doAuth=()=>{try{const cl=google.accounts.oauth2.initTokenClient({client_id:cid.trim(),scope:'https://www.googleapis.com/auth/drive.file',callback:(r)=>{if(r.error||!r.access_token){resolve(null);return;}_driveToken=r.access_token;_driveTokenExpiry=Date.now()+r.expires_in*1000;resolve(_driveToken);}});cl.requestAccessToken({prompt:''});}catch(e){resolve(null);}};
    if(window.google&&google.accounts){doAuth();}else{const s=document.createElement('script');s.src='https://accounts.google.com/gsi/client';s.onload=doAuth;s.onerror=()=>resolve(null);document.head.appendChild(s);}
  });
}
async function _doSalvaSuDrive(pdfUrl,nomeFile){
  showToast('Connessione Google Drive...');
  const token=await getDriveToken();
  if(!token){showToast('Accesso Drive non riuscito','error');downloadDiretto(pdfUrl,nomeFile);return;}
  showToast('Caricamento su Drive...');
  try{
    const res=await fetch(pdfUrl);const blob=await res.blob();
    const meta=JSON.stringify({name:nomeFile,mimeType:'application/pdf'});
    const form=new FormData();form.append('metadata',new Blob([meta],{type:'application/json'}));form.append('file',new File([blob],nomeFile,{type:'application/pdf'}));
    const up=await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,webViewLink',{method:'POST',headers:{'Authorization':'Bearer '+token},body:form});
    if(!up.ok) throw new Error('HTTP '+up.status);
    const r=await up.json();showToast('Salvato su Google Drive!','success');
    if(r.webViewLink)setTimeout(()=>window.open(r.webViewLink,'_blank'),1500);
  }catch(e){showToast('Errore Drive - scarico localmente','error');downloadDiretto(pdfUrl,nomeFile);}
}


// ===== BACKUP EXPORT / IMPORT =====
async function backupExport(){
  if(!savedForms.length){showToast('Nessuna scheda da esportare','error');return;}
  const nomeDefault='ARETE_backup_'+today();
  const nomeInput=prompt('Nome del file di backup:',nomeDefault);
  if(nomeInput===null) return;
  const nomeJson=(nomeInput.trim()||nomeDefault).replace(/\.json$/i,'').replace(/[<>:"/\\|?*]/g,'_')+'.json';
  const payload={version:1,exportedAt:new Date().toISOString(),schede:savedForms};
  const contenuto=JSON.stringify(payload,null,2);
  const blob=new Blob([contenuto],{type:'application/json;charset=utf-8'});
  if(navigator.share&&navigator.canShare){
    try{
      const file=new File([blob],nomeJson,{type:'application/json'});
      if(navigator.canShare({files:[file]})){
        await navigator.share({title:'Backup ARETE',text:'Backup ARETE - '+savedForms.length+' schede - '+today(),files:[file]});
        showToast('Backup salvato!','success');return;
      }
    }catch(e){if(e.name==='AbortError')return;}
  }
  const b64=btoa(unescape(encodeURIComponent(contenuto)));
  const a=document.createElement('a');
  a.href='data:application/json;charset=utf-8;base64,'+b64;
  a.download=nomeJson;a.style.display='none';
  document.body.appendChild(a);a.click();
  setTimeout(()=>document.body.removeChild(a),1000);
  showToast('Backup scaricato','success');
}

function backupImport(input){
  const file=input.files[0];if(!file)return;
  const reader=new FileReader();
  reader.onload=function(e){
    try{
      let t=e.target.result.trim();
      if(!t.startsWith('{')&&!t.startsWith('[')){try{t=decodeURIComponent(escape(atob(t)));}catch(ex){}}
      const raw=JSON.parse(t);
      const schede=Array.isArray(raw)?raw:(raw.schede&&Array.isArray(raw.schede)?raw.schede:null);
      if(!schede||!schede.length){showToast('File non valido o vuoto','error');return;}
      let agg=0,upd=0;
      schede.forEach(imp=>{
        if(!imp.id||!imp.type||!imp.data)return;
        const idx=savedForms.findIndex(f=>f.id===imp.id);
        if(idx<0){savedForms.push(imp);agg++;}
        else if(imp.savedAt>savedForms[idx].savedAt){savedForms[idx]=imp;upd++;}
      });
      localStorage.setItem('arete_forms',JSON.stringify(savedForms));
      renderHomeSaved();renderArchive();
      showToast('Importate: '+agg+' nuove, '+upd+' aggiornate','success');
    }catch(err){showToast('Errore lettura file','error');}
    input.value='';
  };
  reader.readAsText(file,'utf-8');
}

renderHomeSaved();
