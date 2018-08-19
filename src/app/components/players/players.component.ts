import { Component, OnInit, Input, Output, EventEmitter } from '@angular/core';


@Component({
  selector: 'app-players',
  templateUrl: './players.component.html',
  styleUrls: ['./players.component.css']
})
export class PlayersComponent implements OnInit {
  /*players: object =
    {"trevsnan":643.8, "leepops":519.9, "Markbr":479.24, "julien":463.4, "kazbar54":349, "jonkat":333.45, "dean35":310.3, "june1st":267.975, "badbaba":257.925, "bigmags28":248.35, "icklefluer":192.025, "FRAN2312":155.415, "summer522":134.002, "peachyhun":128.905, "JoWilli":125.52, "luckyaaa":118.22, "Terryferry":114, "asp86":99.92, "adoc40":98.275, "Hun123":96.56, "yogebear":96.04, "emmajb19_dor":90.15, "djjmary":89.4, "hannah0908":86.3, "jenniaca04":81.35, "russellh25":80.63, "debbieohyes":79.27, "traceyallen":75.38, "bigmigs1":73.95, "joe0510":73.5, "joecourts":72.315, "basiyasmin":71.325, "saydon":66.5, "sarahk35":64.375, "welfinisawinner":63.37, "gemma71":62.9, "trudy1976":61.978, "Mumma2Be":54.925, "Mama20016":52.975, "Elk89":52.54, "truilee":48.88, "oxokarenoxo":46.5, "kitykat1227":46.39, "stressed":46.3, "Sam679":45.275, "Gardenlady123":44.725, "itscourtneeey":43.58, "laurenpiper88":42.15, "andyt2":40.925, "diffa":40, "EmmaS198322":39.75, "jmoblackbone":38.325, "natsy":37.725, "XNantoJackX":36.68, "olsen04":36.35, "Box":36.1, "ellabellaxx6":35.75, "weetinadoll":34.77, "Tazz0809":33.45, "xXx_Shell_xXx":33.1, "elsie222":31.95, "kirtie":31.75, "kaynruby22":31, "beckiboo1980":30, "mozza1984":29.6, "welsh1972":29.375, "FoxyAli1962":28.9, "izzy25":27.15, "leanneappy85":26.4, "chellh":25.66, "webqueen63":25.5, "Angela":25.3, "carlnshell":24.9, "wendy0810":24.825, "Genna10":24.68, "luckylucifer":24.47, "gailj8":23.975, "vickyk1978":23.94, "rachel60":23.06, "Mrsmillsy":22.84, "cp9792":22.2, "taz1810":21.355, "tonywallen":21.11, "natz88315":21.102, "CC_116":21.09, "imnotinflatable":20.78, "cardiff_fan":20.145, "ME163":19.7, "vicknchris":19.35, "maya94":18.98, "kedgewedge":18.96, "emmatnewcastle":18.7, "rukhsana1":18.25, "jodie0206":17.425, "arnopg":17.4, "bigwrighty67":17.3, "demi67":16.9, "tylercerise64":16.625, "lisa0511":16.525, "FBkellyellis":16.3, "jenjoy2":16.3, "bingostash":16.125, "Bowie123":16, "mazza33":15.6, "sandy4matt":15.575, "michellespiers":15.462, "Nicnac68":15.35, "zaction":15.35, "Louisa1977":15.325, "cc11":15.005, "Steg1418":14.72, "rob39":14.7, "carer":14.625, "Pollycoddle":14.55, "daveamison":14.525, "winner831":14.3, "melliker":14.24, "lizzy1234":14.032, "amanda106":13.7, "Butler2818":13.26, "rarajane1988":13.16, "hillymay":13.15, "chazzygirl":13, "kildare40":12.9, "vilma1":12.825, "lisasin1983":12.7, "kayleigh0795":12.65, "conmisbam":12.6, "birdtown":12.56, "Thena666":12.2, "misskenny":11.975, "steph2610":11.64, "melanie0909":11.44, "orwilaur":11.27, "Kimba20":11, "mariah69":10.855, "ogkusssh":10.825, "pinchy01":10.73, "susie1960":10.7, "chic1985":10.62, "feeley":10.525, "hannahj94":10.35, "gdc":10.25, "tulin11":10.2, "missym1":10.175, "sandishell25":10.14, "simsam33":10.05, "Fraserburgh":10.04, "jata":9.825, "Dr_Evil_1986":9.6, "Yayiwon15":9.465, "Johnny5cott":9.45, "Rscott123456":9.44, "lippy72":9.325, "georgellie":9.125, "ashermay":9.04, "luckylinz":8.61, "maryod":8.6, "wayne1373":8.4, "angela274":8.2, "sandra1986":8.14, "hay93":7.995, "nat35":7.925, "Josie1933":7.905, "michala76":7.875, "carolined30":7.84, "saucysheila19":7.625, "wenjack":7.6, "dawn2514":7.55, "abbielou1997":7.54, "adammaddern":7.5, "Nikes22":7.5, "rippers1":7.5, "ritamaca":7.467, "Reeder86":7.4, "minkyadams":7.39, "steph300412":7.265, "kinderegg":7.185, "tony39":7.175, "stuboy1978":7.15, "lunty08":7.14, "ianr":7.075, "grandmasonic":6.9, "maximus1":6.85, "serava18":6.85, "sandywon":6.825, "BekkieBoo":6.74, "chick1963":6.72, "Paxton77":6.65, "staceyrae13":6.64, "ducks75":6.55, "plantone":6.525, "myles1993":6.5, "rulo2838":6.5, "Lisa191077":6.475, "chloepeel123":6.325, "sacha72":6.21, "witty2018":6.18, "Skippyfromaus":6.14, "KEEPFITMATHESON":6.1, "tracey0105":5.95, "cloeobrien15":5.9, "foxiee1":5.9, "joelyanne":5.85, "danielle1406":5.82, "zsste2":5.755, "finchy01227":5.7, "kirstyrob":5.6, "lilmissduffz":5.58, "bambi1631":5.55, "bigeyes":5.55, "deadeye68":5.5, "Sianee24":5.5, "akettlety":5.44, "crfmx":5.4, "narkycow":5.38, "INBETWEEN":5.3, "JodieSky":5.275, "gsesay27":5.26, "nickola1989":5.2, "BettyBeBe":5.175, "David210":5.15, "dean0707":5.125, "karemem":5.125, "bernieswife":5.1, "terrific":5.06, "rain25":5.05, "bingokid73":5, "hallett7583":4.98, "mydbom":4.955, "sunburst":4.95, "stace2303":4.94, "jakyru":4.92, "yummymummy2b":4.76, "mccollm16":4.74, "jeneratorx":4.73, "alex81100":4.675, "elcally":4.65, "lloydcymru":4.62, "flojojo":4.58, "Jackfrape":4.56, "cpalmer73":4.525, "stephanie107":4.525, "samone1973":4.505, "Vixster41":4.5, "bingogirl9":4.45, "bowe1956":4.425, "tbevis92":4.42, "kenneth81":4.395, "shilasmammy":4.35, "yesplz2":4.337, "twesty":4.28, "nicguy31":4.275, "Nuxe1959":4.2, "meggymoomoo":4.195, "oliviosa":4.18, "samj2018":4.16, "LEAH579":4.15, "leosmummy141":4.15, "chelle2111":4.125, "ZANNY":4.125, "beth28384":4.1, "duffmar08":4.1, "Lady13":4.1, "Loopy5543":4.08, "maximumsoot":4.04, "MITCH1981_dor":4.025, "aymlou18":4.02, "hay123":4, "lambrettagirl":3.95, "parkside158":3.94, "raffers92":3.94, "charlanc1990":3.9, "gtsbabyxo":3.88, "Carry":3.875, "Nickier":3.875, "sandyjayne":3.87, "mickygrieves":3.86, "elma1983":3.85, "tay07":3.85, "Emilyx86":3.825, "kayler1114":3.82, "Malibumum":3.8, "celticlads":3.75, "samriley0201":3.75, "ievute":3.74, "tiffanyprincess":3.74, "charley99":3.7, "kerryleather":3.7, "Malachi10":3.68, "vw831":3.68, "alton":3.675, "susiexxx":3.675, "minkyboo":3.66, "janeatkins1973":3.64, "dinocat55":3.6, "chelmac2018":3.58, "Kazzyjane":3.55, "grandadmo":3.54, "suzzliz":3.5, "keelreal":3.46, "chloenixon1993":3.42, "elizabeth83":3.42, "atesz11201":3.4, "turkishwifey":3.4, "acabbieswife":3.35, "Baggins11":3.35, "maria4awin":3.35, "madameshabooboo":3.34, "LABBY08":3.32, "paulevans66":3.3, "stepheywins":3.3, "Tomatoplants":3.3, "jamiesmum":3.25, "sheana86":3.21, "annh18":3.2, "chloe135":3.2, "lorrain9":3.2, "Neiceynoo":3.2, "Zoey1211":3.2, "kathmac81":3.15, "stebbos":3.15, "falcor":3.14, "Birch12_dor":3.1, "buff22":3.1, "G93Bailey":3.1, "Nik723":3.1, "tara6667":3.1, "babygee197":3.075, "hayleylovescarl":3.075, "gilly1966":3.025, "rara24683":3.02, "gemsie1984":3, "nannyjan56":3, "ste105":3, "kevin26":2.98, "petesellie":2.95, "lillypants2016":2.925, "weeman35":2.925, "Chelseaxsky":2.92, "sikara01":2.904, "duckster":2.9, "amanda2401":2.88, "simon38":2.82, "yvonne11":2.82, "missvenus":2.8, "Sherbelli":2.8, "angiez":2.775, "JustKenny":2.75, "dazandali":2.725, "lgilsie":2.725, "bubbzbaybee7":2.72, "chaz7598":2.71, "leezee01":2.7, "mojo94":2.7, "Gertrude01":2.675, "Daftdolly":2.65, "inthegarden_dor":2.65, "hickey18":2.625, "migetgem45":2.625, "Majorcaticket":2.57, "halifaxjo":2.56, "hayz8186":2.5, "lila1":2.5, "omg111":2.5, "SnowWhite26_dor":2.5, "vas2018":2.48, "whatsinit4me":2.476, "momo1939":2.455, "charlieboy5":2.415, "ace75_dor":2.4, "babyj":2.4, "jadeycoomber":2.4, "robgenmarjen":2.4, "rose1neverwin":2.4, "shapstaf":2.4, "tonijew":2.375, "Duckduckgoose":2.36, "Snowfy":2.36, "samgirl23":2.35, "Vanillapodx":2.34, "hannahmorning":2.32, "alisonm1979":2.3, "GRUNNIE05":2.3, "joey0881":2.3, "karen100":2.3, "Solo1974":2.3, "wiin4me":2.3, "lilypie1978":2.26, "jeanetta":2.25, "mishiebrown":2.245, "Paigeisobel1":2.24, "vampsx":2.24, "bonkas0":2.2, "dirdyhelen24":2.2, "beau1952":2.15, "daz4040":2.15, "katmason12":2.1, "Tiona64":2.1, "patsyann1951":2.08, "cruisecrazy":2.075, "salt21":2.075, "seale":2.06, "stakesby":2.025, "dionne1981":2, "Madmandyms45":2, "Seanybeast":2, "laurenk195":1.95, "luckycarol4":1.945, "Icekaren":1.9, "WhoopWhoopWinna":1.9, "frooty6":1.88, "kegalish36":1.86, "Suritai2":1.86, "hayden10":1.85, "Jods3523":1.85, "billywizz2000":1.84, "Skyela99":1.84, "makemelucky5":1.825, "lupinda":1.823, "carter690":1.82, "mummy1980s":1.82, "nannytroy":1.8, "sassy04":1.8, "agents":1.78, "Kev0406":1.78, "Zieziej":1.78, "angel6060":1.775, "rjenkins34":1.775, "wongy74":1.775, "kimbo1981":1.76, "gran2jrc":1.75, "lilkaz":1.75, "xxxlogi":1.75, "annrodger":1.725, "Auchenflower":1.725, "ChloeAnn95":1.725, "Johnsgirl60":1.725, "junkiedred":1.725, "ebony5":1.72, "discountdave":1.7, "martine45":1.7, "Gemlibobs":1.68, "keith1210":1.68, "angietee":1.65, "buzzwuzz":1.65, "debsie76":1.65, "minnie1":1.65, "loveangel77":1.64, "bingoqueen5":1.629, "voila":1.625, "andy070707":1.6, "paneho":1.6, "Sarah72p":1.6, "skintflintcarer":1.6, "adnil21":1.575, "shelly2005":1.575, "fishermangold":1.54, "richard5":1.5, "suzzzey":1.5, "xokimboxo77":1.5, "kenneth2803":1.483, "cobblers4":1.48, "Meglayla13":1.475, "smeg0712":1.475, "Jessiewins":1.46, "liamlod890":1.45, "kimncarl":1.44, "tigger1231":1.44, "sherob2121":1.425, "leaandyxxx":1.42, "charliecraft":1.4, "luckyjake123":1.4, "Pmxxxx":1.4, "suziesoo63":1.4, "tonid17":1.385, "tinasluck":1.38, "WORMING":1.36, "wolfcharmed":1.355, "Amyd891":1.354, "kirsty2357":1.35, "nicnic1":1.35, "bigdadjohn":1.34, "willie49":1.336, "nikkip79":1.335, "Rudegirl40":1.33, "flutter34":1.325, "lindsay92":1.31, "bexta24":1.3, "darost":1.3, "newlife2018":1.3, "princesspwills":1.3, "y33uth":1.3, "lisa6":1.275, "fleur0069":1.26, "Lainey123_dor":1.25, "nikkinoo13":1.24, "Di1980":1.225, "aaronwinner1":1.22, "kirbs1":1.22, "ew1702":1.21, "boow_dor":1.2, "flowerpower5":1.2, "johneyc":1.2, "palace68":1.2, "suzie72":1.2, "cathhead":1.18, "lindsey1810":1.18, "suejp63":1.18, "teesha":1.18, "lollipop129":1.16, "samzbamz":1.16, "ttar":1.16, "nannapat1":1.15, "rachbaby1":1.15, "natalie61":1.14, "JAZMIN1993":1.125, "superscot8":1.125, "lesleyf82":1.12, "Dave180586":1.1, "emmrich":1.1, "Kattrix1":1.1, "Kimmilimbo":1.1, "staceb1":1.1, "vikki1976":1.1, "Njackson87":1.09, "alf1983":1.08, "mayflower53":1.08, "mrwoofles5":1.08, "tracinpaper":1.08, "annie6446":1.04, "jbro221":1.04, "jshan2":1.04, "mythreebabys1":1.04, "shazbo2":1.04, "Chelseakate96":1.035, "danilinzy":1.03, "rosarocco":1.025, "mandydawl":1.02, "Rafakins":1.02, "beckystar111":1, "cutiexlou":1, "kirstywilliamsl":1, "layla59":1, "lynny1958":1, "oxon":1, "WoodysMam13":1} 
  ;*/
  @Input() players:[String, number][];
  @Input() export: Function;
  @Output() messageEvent = new EventEmitter<[String, number, number][]>();

  playersAsArrays:[String, number][];
  range:[String, number, number][];
  subTotal: number;
  numDraw: number;
  winnerArr: [String, number, number][];
  oneWin: boolean;

  constructor() { }

  ngOnInit() {
   
  }

  sendMessage() {
    this.messageEvent.emit(this.winnerArr)
  }
  
  getPlayers():void{
    this.playersAsArrays = this.players; 
    console.log("getPlayers");
    
    
    this.getRange();
    this.getWinner();
    
  }

  getRange():void{
    //console.log("getRange");
    this.range=new Array();
	
    this.playersAsArrays.map( (element , index)=> {
      let subTotal=0;
      for (let i=0; i<=index; i++){
        
        //first item of the range element is player name
        if (!this.range[i]){
          this.range[i] = [element[0],0,0];
        }

        if (!this.range[i][0]){
          this.range[i][0] = element[0];
        }
        
        //second item of the range element is the start number of this player's range
        this.range[i][1]=subTotal + 1;
        subTotal += Math.round(this.playersAsArrays[i][1]);

        //third item of the range element is the end number of this player's range
        this.range[i][2]=subTotal;
        
      }
    
      this.subTotal = subTotal;
		
	});
	  //console.log("subTotal ", this.subTotal);
  }

  getWinner() {
    this.winnerArr=new Array();
    
    do {
      
      let rdm = Math.floor(Math.random() * this.subTotal) + 1  
    
      let winner = this.range.filter(entryRange => ( rdm >= entryRange[1] && rdm <= entryRange[2]));
     
      //console.log("winner " + winner[0]);
      
      if (!this.oneWin) {
        winner[0].splice(3,0,rdm);
        this.winnerArr.push(winner[0]);
      } else {
        if (!this.winnerArr.includes(winner[0])){
          winner[0].splice(3,0,rdm);
          this.winnerArr.push(winner[0])
        }
      }
 
      
    } while (this.winnerArr.length < this.numDraw);
    
    //console.log("xxxx getWinner end winnerArr= " + this.winnerArr);
    this.sendMessage();
  }

  

}