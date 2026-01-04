// =============================================================================
// ENHANCED EXCEL TO EBAY CONVERTER - V2 WITH ALL FIXES
// =============================================================================

document.addEventListener("DOMContentLoaded", () => {
  let excelData = [];
  let pdfDataArray = [];
  let categoryMap = null;

  // =============================================================================
  // PRODUCT LOOKUP - DISABLED (API calls removed for now)
  // =============================================================================

  // API-based product lookup has been disabled to avoid rate limits.
  // The category matching will use title-based matching instead.
  // To re-enable in future, add Google Custom Search API integration here.

  const STOPWORDS = new Set([
    "with",
    "for",
    "and",
    "or",
    "the",
    "a",
    "an",
    "in",
    "on",
    "at",
    "to",
    "from",
    "by",
    "of",
    "is",
    "was",
    "are",
    "were",
    "been",
    "this",
    "that",
    "these",
    "those",
    "it",
    "its",
    "set",
    "pack",
    "pcs",
    "stock",
    "2019",
    "2020",
    "2021",
    "2022",
    "2023",
    "2024",
    "2025",
    "cm",
    "mm",
    "inch",
    "inches",
  ]);

  // =============================================================================
  // DESCRIPTION GENERATOR CLASS
  // =============================================================================

  class DescriptionGenerator {
    constructor() {
      this.templates = {
        Pillow: (brand, title, specs) =>
          `${
            brand ? brand + " " : ""
          }${title} - Premium quality pillow for comfort and support. ${specs}Perfect for a great night's sleep.`,
        "Memory Foam Pillow": (brand, title, specs) =>
          `${
            brand ? brand + " " : ""
          }${title} - Memory foam pillow with excellent neck support. ${specs}Ideal for comfortable sleep.`,
        "Phone Case": (brand, title, specs) =>
          `${
            brand ? brand + " " : ""
          }${title} - Protective phone case with premium design. ${specs}Keep your device safe and stylish.`,
        "Tablet Case": (brand, title, specs) =>
          `${
            brand ? brand + " " : ""
          }${title} - Tablet case for protection and style. ${specs}Keep your device safe and functional.`,
        "iPad Case": (brand, title, specs) =>
          `${
            brand ? brand + " " : ""
          }${title} - iPad case for protection and style. ${specs}Keep your device safe and functional.`,
        "Screen Protector": (brand, title, specs) =>
          `${
            brand ? brand + " " : ""
          }${title} - Screen protector for device protection. ${specs}Crystal clear and scratch resistant.`,
        "Dog Grooming": (brand, title, specs) =>
          `${
            brand ? brand + " " : ""
          }${title} - Professional pet grooming tool. ${specs}Keep your pet looking great.`,
        "Pet Clippers": (brand, title, specs) =>
          `${
            brand ? brand + " " : ""
          }${title} - Quality pet clippers for grooming. ${specs}Easy to use and effective.`,
        Curtains: (brand, title, specs) =>
          `${
            brand ? brand + " " : ""
          }${title} - Quality curtains for your home. ${specs}Enhance your living space with style.`,
        Fan: (brand, title, specs) =>
          `${
            brand ? brand + " " : ""
          }${title} - Cooling fan for home or office. ${specs}Stay cool and comfortable.`,
        "Air Mattress": (brand, title, specs) =>
          `${
            brand ? brand + " " : ""
          }${title} - Inflatable air mattress for comfort. ${specs}Perfect for guests or camping.`,
        "Bed Risers": (brand, title, specs) =>
          `${
            brand ? brand + " " : ""
          }${title} - Furniture risers for extra height. ${specs}Sturdy and reliable support.`,
        Electronics: (brand, title, specs) =>
          `${
            brand ? brand + " " : ""
          }${title} - Premium electronic device. ${specs}Perfect for modern use.`,
        Footwear: (brand, title, specs) =>
          `${
            brand ? brand + " " : ""
          }${title} - Comfortable and stylish footwear. ${specs}Perfect for any occasion.`,
        Clothing: (brand, title, specs) =>
          `${
            brand ? brand + " " : ""
          }${title} - Quality apparel for style and comfort. ${specs}Versatile for various occasions.`,
        Home: (brand, title, specs) =>
          `${
            brand ? brand + " " : ""
          }${title} - Quality home product. ${specs}Enhance your living space.`,
        Kitchen: (brand, title, specs) =>
          `${
            brand ? brand + " " : ""
          }${title} - Essential kitchen item. ${specs}Designed for everyday use.`,
        Beauty: (brand, title, specs) =>
          `${
            brand ? brand + " " : ""
          }${title} - Quality beauty product. ${specs}Professional grade formulation.`,
        Toys: (brand, title, specs) =>
          `${
            brand ? brand + " " : ""
          }${title} - Quality toy for fun and learning. ${specs}Great for all ages.`,
        Pet: (brand, title, specs) =>
          `${
            brand ? brand + " " : ""
          }${title} - Quality pet product. ${specs}Keep your pet happy and healthy.`,
        Garden: (brand, title, specs) =>
          `${
            brand ? brand + " " : ""
          }${title} - Garden and outdoor product. ${specs}Perfect for your outdoor space.`,
        Sports: (brand, title, specs) =>
          `${
            brand ? brand + " " : ""
          }${title} - Sports equipment for performance. ${specs}Designed for athletes.`,
        Baby: (brand, title, specs) =>
          `${
            brand ? brand + " " : ""
          }${title} - Quality baby product. ${specs}Safe for your little one.`,
        Automotive: (brand, title, specs) =>
          `${
            brand ? brand + " " : ""
          }${title} - Automotive accessory. ${specs}Keep your vehicle in top condition.`,
        default: (brand, title, specs) =>
          `${
            brand ? brand + " " : ""
          }${title} - Premium quality product. ${specs}Excellent choice for your needs.`,
      };
    }

    generateDescription(title, brand, itemType, specifics) {
      const specificsText = this.buildSpecificsText(specifics);
      const template = this.templates[itemType] || this.templates.default;
      const description = template(brand, title, specificsText);
      return description.substring(0, 500);
    }

    buildSpecificsText(specifics) {
      if (!specifics || Object.keys(specifics).length === 0) return "";
      const features = [];
      for (const [key, value] of Object.entries(specifics)) {
        if (value && value !== "" && key !== "Brand" && key !== "Type") {
          features.push(`${key}: ${value}`);
        }
      }
      if (features.length === 0) return "";
      return "Features: " + features.slice(0, 5).join(", ") + ". ";
    }
  }

  const descriptionGenerator = new DescriptionGenerator();

  // =============================================================================
  // ENHANCED SPECIFICS GENERATOR - WITH MORE SPECIFICS
  // =============================================================================

  class EnhancedSpecificsGenerator {
    constructor() {
      // Brand patterns
      this.brands = {
        Apple:
          /\b(apple|iphone|ipad|macbook|airpods|imac|ipod|apple\s*watch)\b/i,
        Samsung: /\b(samsung|galaxy|note\s*\d|tab\s*[as])\b/i,
        Sony: /\b(sony|playstation|ps[45]|bravia|xperia)\b/i,
        Microsoft: /\b(microsoft|xbox|surface)\b/i,
        Dell: /\b(dell|alienware|xps|inspiron)\b/i,
        HP: /\b(hp|hewlett|pavilion|omen|envy)\b/i,
        Lenovo: /\b(lenovo|thinkpad|yoga|ideapad)\b/i,
        ASUS: /\b(asus|vivobook|tuf|rog|zenbook)\b/i,
        Nike: /\b(nike|air\s*max|air\s*jordan|air\s*force)\b/i,
        Adidas: /\b(adidas|ultraboost|nmd|yeezy)\b/i,
        Puma: /\b(puma)\b/i,
        "New Balance": /\b(new\s*balance)\b/i,
        Canon: /\b(canon|eos|powershot)\b/i,
        Nikon: /\b(nikon|coolpix)\b/i,
        Dyson: /\b(dyson)\b/i,
        Philips: /\b(philips|sonicare)\b/i,
        LG: /\b(lg|oled|nanocell)\b/i,
        Bosch: /\b(bosch)\b/i,
        Makita: /\b(makita)\b/i,
        DeWalt: /\b(dewalt)\b/i,
        Lego: /\b(lego)\b/i,
        Nintendo: /\b(nintendo|switch|wii)\b/i,
        JBL: /\b(jbl)\b/i,
        Bose: /\b(bose)\b/i,
        Beats: /\b(beats)\b/i,
        Logitech: /\b(logitech)\b/i,
        Razer: /\b(razer)\b/i,
        KitchenAid: /\b(kitchenaid)\b/i,
        Ninja: /\b(ninja)\b/i,
        Tefal: /\b(tefal|t-fal)\b/i,
        Breville: /\b(breville)\b/i,
        NOFFA: /\b(noffa)\b/i,
        SEYMAC: /\b(seymac)\b/i,
        Doeshine: /\b(doeshine)\b/i,
        Marchpower: /\b(marchpower)\b/i,
        iPrimio: /\b(iprimio)\b/i,
        OPICK: /\b(opick)\b/i,
      };

      // Item Type detection patterns - returns SPECIFIC product type
      this.typePatterns = [
        // Bedding
        { pattern: /\bmemory\s*foam\s*pillow\b/i, type: "Memory Foam Pillow" },
        { pattern: /\bpillow\b/i, type: "Pillow" },
        { pattern: /\bduvet\b/i, type: "Duvet" },
        { pattern: /\bblanket\b/i, type: "Blanket" },
        { pattern: /\bcushion\b/i, type: "Cushion" },
        { pattern: /\bmattress\s*(topper|pad)?\b/i, type: "Mattress" },
        { pattern: /\bbed\s*sheet/i, type: "Bed Sheet" },

        // Cases & Protectors
        { pattern: /\bipad\s*(case|cover)\b/i, type: "iPad Case" },
        { pattern: /\btablet\s*(case|cover)\b/i, type: "Tablet Case" },
        {
          pattern: /\b(phone|iphone|samsung|galaxy)\s*(case|cover)\b/i,
          type: "Phone Case",
        },
        { pattern: /\bscreen\s*protect/i, type: "Screen Protector" },
        { pattern: /\bcase\b.*\b(ipad|tablet)\b/i, type: "Tablet Case" },
        {
          pattern: /\bcase\b.*\b(phone|iphone|samsung|galaxy)\b/i,
          type: "Phone Case",
        },

        // Pet Products
        {
          pattern: /\b(dog|pet)\s*(clipper|trimmer|grooming)/i,
          type: "Pet Clippers",
        },
        {
          pattern: /\bdog\s*(bed|collar|leash|toy|bowl|food)/i,
          type: "Dog Supplies",
        },
        {
          pattern: /\bcat\s*(bed|collar|toy|food|litter)/i,
          type: "Cat Supplies",
        },
        { pattern: /\bpet\s*(bed|toy|bowl|carrier)/i, type: "Pet Supplies" },

        // Medical / Mobility
        { pattern: /\b(knee\s*)?crutch/i, type: "Crutch" },
        { pattern: /\bmobility\s*aid/i, type: "Mobility Aid" },
        { pattern: /\bwalking\s*(aid|frame|stick)/i, type: "Walking Aid" },
        { pattern: /\bwheelchair/i, type: "Wheelchair" },
        { pattern: /\bknee\s*pad/i, type: "Knee Pad" },

        // Baby Products
        { pattern: /\b(baby\s*)?playpen/i, type: "Playpen" },
        { pattern: /\bhigh\s*chair/i, type: "High Chair" },
        { pattern: /\bhighchair/i, type: "High Chair" },
        { pattern: /\bhigh\s*chair\s*tray/i, type: "High Chair Tray" },
        { pattern: /\bbaby\s*gate/i, type: "Baby Gate" },
        { pattern: /\bbaby\s*monitor/i, type: "Baby Monitor" },
        { pattern: /\bpram\b/i, type: "Pram" },
        { pattern: /\bpushchair/i, type: "Pushchair" },
        { pattern: /\bstroller/i, type: "Stroller" },
        { pattern: /\bcar\s*seat/i, type: "Car Seat" },

        // Kitchen Tools
        { pattern: /\b(cheese\s*)?grater/i, type: "Grater" },
        { pattern: /\b(vegetable\s*)?slicer/i, type: "Slicer" },
        { pattern: /\bmandoline/i, type: "Mandoline" },
        { pattern: /\bpeeler/i, type: "Peeler" },
        { pattern: /\bchopper/i, type: "Chopper" },

        // Air Quality / Appliances
        { pattern: /\bair\s*purifier/i, type: "Air Purifier" },
        { pattern: /\bhepa\s*filter/i, type: "HEPA Filter" },
        { pattern: /\breplacement\s*filter/i, type: "Replacement Filter" },
        { pattern: /\bhumidifier/i, type: "Humidifier" },
        { pattern: /\bdehumidifier/i, type: "Dehumidifier" },

        // Home - Fans
        {
          pattern: /\b(standing|pedestal|tower|desk|portable)\s*fan\b/i,
          type: "Fan",
        },
        { pattern: /\bfan\b/i, type: "Fan" },

        // Home - Curtains & Blinds
        { pattern: /\bdoor\s*curtain/i, type: "Door Curtain" },
        { pattern: /\bthermal\s*curtain/i, type: "Thermal Curtain" },
        { pattern: /\bblackout\s*curtain/i, type: "Blackout Curtain" },
        { pattern: /\bcurtain/i, type: "Curtains" },
        { pattern: /\bblind/i, type: "Blinds" },

        // Furniture
        { pattern: /\bbed\s*riser/i, type: "Bed Risers" },
        { pattern: /\bfurniture\s*riser/i, type: "Furniture Risers" },
        { pattern: /\b(sofa|couch)\b/i, type: "Sofa" },
        { pattern: /\bchair\b/i, type: "Chair" },
        { pattern: /\btable\b/i, type: "Table" },
        { pattern: /\bdesk\b/i, type: "Desk" },

        // Air Mattress
        { pattern: /\b(inflatable|air)\s*mattress\b/i, type: "Air Mattress" },
        { pattern: /\bair\s*bed\b/i, type: "Air Mattress" },

        // Electronics
        {
          pattern: /\b(headphone|earphone|earbud|airpod)/i,
          type: "Headphones",
        },
        { pattern: /\bspeaker/i, type: "Speaker" },
        { pattern: /\bcharger\b/i, type: "Charger" },
        { pattern: /\bkeyboard\b/i, type: "Keyboard" },
        { pattern: /\bmouse\b/i, type: "Mouse" },
        { pattern: /\bmonitor\b/i, type: "Monitor" },
        { pattern: /\btv\b|\btelevision\b/i, type: "Television" },
        { pattern: /\bcamera\b/i, type: "Camera" },
        { pattern: /\blaptop\b/i, type: "Laptop" },
        { pattern: /\btablet\b/i, type: "Tablet" },
        { pattern: /\bwatch\b/i, type: "Watch" },
        { pattern: /\bcd\s*player/i, type: "CD Player" },
        { pattern: /\bdiscman/i, type: "CD Player" },
        { pattern: /\bportable\s*cd/i, type: "CD Player" },

        // Gaming
        {
          pattern: /\b(xbox|playstation|ps[\-]?[45])\s*controller/i,
          type: "Game Controller",
        },
        { pattern: /\bgamepad/i, type: "Game Controller" },
        { pattern: /\bwired\s*controller/i, type: "Game Controller" },
        { pattern: /\bgaming\s*controller/i, type: "Game Controller" },

        // Memorial / Urns
        { pattern: /\burn\b.*\bash/i, type: "Cremation Urn" },
        { pattern: /\bcremation\s*urn/i, type: "Cremation Urn" },
        { pattern: /\burn\b/i, type: "Urn" },
        { pattern: /\bpet\s*memorial/i, type: "Pet Memorial Urn" },

        // Barn Door / Door Hardware
        {
          pattern: /\bbarn\s*door\s*(hardware|kit)/i,
          type: "Barn Door Hardware",
        },
        { pattern: /\bsliding\s*barn\s*door/i, type: "Barn Door Hardware" },
        { pattern: /\bdouble\s*barn\s*door/i, type: "Barn Door Hardware" },
        {
          pattern: /\bsliding\s*door\s*(hardware|kit)/i,
          type: "Sliding Door Hardware",
        },
        { pattern: /\bdoor\s*track/i, type: "Door Hardware" },
        { pattern: /\bceiling\s*mount.*door/i, type: "Barn Door Hardware" },

        // Baby Play Mat
        { pattern: /\bbaby\s*play\s*mat/i, type: "Baby Play Mat" },
        { pattern: /\bfoam\s*play\s*mat/i, type: "Foam Play Mat" },
        { pattern: /\bplaymat/i, type: "Play Mat" },
        { pattern: /\bplay\s*mat/i, type: "Play Mat" },
        { pattern: /\bcrawling\s*mat/i, type: "Crawling Mat" },
        { pattern: /\bfloor\s*mat.*baby/i, type: "Baby Floor Mat" },
        { pattern: /\bfoldable\s*mat/i, type: "Foldable Mat" },

        // Caster Wheels
        { pattern: /\bcaster\s*wheel/i, type: "Caster Wheels" },
        { pattern: /\bcasters/i, type: "Casters" },
        { pattern: /\bheavy\s*duty\s*caster/i, type: "Heavy Duty Casters" },
        { pattern: /\bindustrial\s*caster/i, type: "Industrial Casters" },

        // Decorative Box
        { pattern: /\bdecorative\s*box/i, type: "Decorative Box" },
        { pattern: /\bglass.*lidded\s*box/i, type: "Glass Lidded Box" },
        { pattern: /\btrinket\s*box/i, type: "Trinket Box" },
        { pattern: /\bvintage.*box/i, type: "Vintage Box" },

        // Drain Rods
        { pattern: /\bdrain\s*rod/i, type: "Drain Rods" },
        { pattern: /\bdrain\s*cleaning/i, type: "Drain Cleaning Equipment" },

        // Footmuff / Cosytoes
        { pattern: /\bcosytoes/i, type: "Cosytoes" },
        { pattern: /\bfootmuff/i, type: "Footmuff" },
        { pattern: /\bbaby\s*footmuff/i, type: "Baby Footmuff" },
        { pattern: /\bhandmuff/i, type: "Handmuff" },

        // Galaxy Z Fold Case
        { pattern: /\bgalaxy\s*z\s*fold.*case/i, type: "Galaxy Z Fold Case" },
        { pattern: /\bz\s*fold.*case/i, type: "Z Fold Case" },
        { pattern: /\bsamsung.*case/i, type: "Samsung Phone Case" },

        // Plumbing
        { pattern: /\bcistern/i, type: "Cistern" },
        { pattern: /\btoilet/i, type: "Toilet" },

        // Garment/Luggage
        { pattern: /\bsuit\s*bag/i, type: "Suit Bag" },
        { pattern: /\bgarment\s*bag/i, type: "Garment Bag" },

        // Garden/Outdoor
        { pattern: /\btable\s*cover/i, type: "Furniture Cover" },
        { pattern: /\bgarden.*cover/i, type: "Garden Furniture Cover" },
        { pattern: /\bpatio.*cover/i, type: "Patio Furniture Cover" },
        { pattern: /\bhot\s*tub\s*cover/i, type: "Hot Tub Cover" },
        { pattern: /\bhot\s*tub/i, type: "Hot Tub" },

        // Pet Training
        { pattern: /\bbark\s*(control|collar)/i, type: "Bark Control Collar" },
        { pattern: /\banti\s*bark/i, type: "Bark Control Device" },
        { pattern: /\bdog\s*steps/i, type: "Dog Steps" },
        { pattern: /\bpet\s*stairs/i, type: "Pet Stairs" },
        {
          pattern: /\bcat\s*(repellent|scarer|deterrent)/i,
          type: "Cat Repellent",
        },
        { pattern: /\bpest\s*repeller/i, type: "Pest Repeller" },
        { pattern: /\bultrasonic.*repel/i, type: "Ultrasonic Repeller" },

        // Hand Dryer
        { pattern: /\bhand\s*dryer/i, type: "Hand Dryer" },
        {
          pattern: /\bcommercial\s*hand\s*dryer/i,
          type: "Commercial Hand Dryer",
        },

        // Bluetooth / Wireless Headset
        { pattern: /\bbluetooth\s*headset/i, type: "Bluetooth Headset" },
        { pattern: /\bwireless\s*headset/i, type: "Wireless Headset" },
        { pattern: /\bnoise\s*cancell/i, type: "Noise Cancelling Headset" },

        // Hook-on High Chair
        { pattern: /\bhook[\s-]?on.*chair/i, type: "Hook-on High Chair" },
        { pattern: /\bportable\s*high\s*chair/i, type: "Portable High Chair" },

        // Thermostat / Heating
        { pattern: /\bwifi\s*thermostat/i, type: "WiFi Thermostat" },
        { pattern: /\bsmart\s*thermostat/i, type: "Smart Thermostat" },
        { pattern: /\bthermostat/i, type: "Thermostat" },
        { pattern: /\bunderfloor\s*heating/i, type: "Underfloor Heating" },
        { pattern: /\bfloor\s*heating/i, type: "Floor Heating" },

        // TV Remote
        { pattern: /\btv\s*remote/i, type: "TV Remote" },
        { pattern: /\bsmart\s*tv\s*remote/i, type: "Smart TV Remote" },
        { pattern: /\breplacement\s*remote/i, type: "Replacement Remote" },

        // Passport Cover / Travel
        { pattern: /\bpassport\s*cover/i, type: "Passport Cover" },
        { pattern: /\bpassport\s*holder/i, type: "Passport Holder" },
        { pattern: /\bpassport\s*case/i, type: "Passport Case" },

        // Party Supplies
        { pattern: /\bplastic\s*cups/i, type: "Plastic Cups" },
        { pattern: /\bparty\s*cups/i, type: "Party Cups" },
        { pattern: /\btumblers/i, type: "Tumblers" },

        // Humidity Control
        { pattern: /\bhumidity\s*controller/i, type: "Humidity Controller" },
        { pattern: /\bhumidistat/i, type: "Humidistat" },
        { pattern: /\bthermostat/i, type: "Thermostat" },

        // Clothing & Footwear
        { pattern: /\bshorts\b/i, type: "Shorts" },
        { pattern: /\blinen\s*shorts/i, type: "Linen Shorts" },
        { pattern: /\b(trainer|sneaker|running\s*shoe)/i, type: "Trainers" },
        { pattern: /\bboot/i, type: "Boots" },
        { pattern: /\bsandal/i, type: "Sandals" },
        { pattern: /\bshoe/i, type: "Shoes" },
        { pattern: /\b(t-shirt|tshirt|tee)\b/i, type: "T-Shirt" },
        { pattern: /\bshirt\b/i, type: "Shirt" },
        { pattern: /\bdress\b/i, type: "Dress" },
        { pattern: /\b(trouser|pant|jean)/i, type: "Trousers" },
        { pattern: /\bjacket\b/i, type: "Jacket" },
        { pattern: /\bcoat\b/i, type: "Coat" },
        { pattern: /\bhoodie\b/i, type: "Hoodie" },
        { pattern: /\bsweater\b/i, type: "Sweater" },

        // Kitchen
        { pattern: /\b(pan|pot|skillet)\b/i, type: "Cookware" },
        { pattern: /\bknife\b|knives\b/i, type: "Knife" },
        { pattern: /\bblender\b/i, type: "Blender" },
        { pattern: /\btoaster\b/i, type: "Toaster" },
        { pattern: /\bkettle\b/i, type: "Kettle" },
        { pattern: /\bcoffee\s*maker\b/i, type: "Coffee Maker" },
        { pattern: /\bpasta\s*maker/i, type: "Pasta Maker" },

        // Office Equipment
        { pattern: /\bpaper\s*shredder/i, type: "Paper Shredder" },
        { pattern: /\bshredder\b/i, type: "Shredder" },

        // Coats / Outerwear
        { pattern: /\btrench\s*coat/i, type: "Trench Coat" },
        { pattern: /\bwool\s*coat/i, type: "Wool Coat" },
        { pattern: /\bovercoat/i, type: "Overcoat" },
        { pattern: /\bcoat\b/i, type: "Coat" },

        // Headsets
        { pattern: /\bpc\s*headset/i, type: "PC Headset" },
        { pattern: /\bbusiness\s*headset/i, type: "Business Headset" },
        { pattern: /\bheadset\b/i, type: "Headset" },

        // Laptop Bags
        { pattern: /\blaptop\s*sleeve/i, type: "Laptop Sleeve" },
        { pattern: /\blaptop\s*case/i, type: "Laptop Case" },
        { pattern: /\blaptop\s*bag/i, type: "Laptop Bag" },

        // Watches
        { pattern: /\bmen.*watch/i, type: "Men's Watch" },
        { pattern: /\bwomen.*watch/i, type: "Women's Watch" },
        { pattern: /\bwristwatch/i, type: "Wristwatch" },
        { pattern: /\bwatch\b/i, type: "Watch" },

        // Rotary Tools
        { pattern: /\brotary\s*tool/i, type: "Rotary Tool" },
        { pattern: /\bmulti\s*tool/i, type: "Multi Tool" },

        // Grooming
        { pattern: /\bbeard\s*trimmer/i, type: "Beard Trimmer" },
        { pattern: /\bstubble\s*trimmer/i, type: "Stubble Trimmer" },

        // Home Decor
        { pattern: /\broom\s*divider/i, type: "Room Divider" },
        { pattern: /\bcurtain\s*rod/i, type: "Curtain Rod" },
        { pattern: /\bblackout\s*curtain/i, type: "Blackout Curtains" },

        // Remote Controls
        { pattern: /\breplacement\s*remote/i, type: "Replacement Remote" },
        { pattern: /\bremote\s*control/i, type: "Remote Control" },

        // Appliance Parts
        { pattern: /\bcutlery\s*basket/i, type: "Cutlery Basket" },
        { pattern: /\bdishwasher.*basket/i, type: "Dishwasher Basket" },
        { pattern: /\boven.*glass/i, type: "Oven Glass Cover" },
        { pattern: /\bdryer\s*vent/i, type: "Dryer Vent Hose" },
        { pattern: /\bvent\s*hose/i, type: "Vent Hose" },

        // Plumbing
        { pattern: /\bfilling\s*valve/i, type: "Filling Valve" },
        { pattern: /\bcistern\s*valve/i, type: "Cistern Valve" },

        // Cassette Player
        { pattern: /\bcassette\s*player/i, type: "Cassette Player" },

        // Medical / Care
        { pattern: /\bbed\s*alarm/i, type: "Bed Alarm" },
        { pattern: /\bcaregiver\s*pager/i, type: "Caregiver Pager" },
        { pattern: /\bmobility\s*aid/i, type: "Mobility Aid" },
        { pattern: /\bstand\s*assist/i, type: "Stand Assist" },

        // Baby
        { pattern: /\bbaby\s*sleeping\s*bag/i, type: "Baby Sleeping Bag" },
        { pattern: /\bsleep\s*sack/i, type: "Sleep Sack" },

        // Water Bottle
        { pattern: /\bwater\s*bottle/i, type: "Water Bottle" },

        // Outdoor Lighting
        { pattern: /\boutdoor.*wall.*light/i, type: "Outdoor Wall Light" },
        { pattern: /\bmotion\s*sensor.*light/i, type: "Motion Sensor Light" },
        { pattern: /\bsecurity\s*light/i, type: "Security Light" },

        // Beauty
        { pattern: /\b(makeup|cosmetic)/i, type: "Makeup" },
        { pattern: /\bperfume\b|\bfragrance\b/i, type: "Perfume" },
        {
          pattern: /\bhair\s*(dryer|straightener|curler)/i,
          type: "Hair Styling",
        },
        { pattern: /\bskincare\b|\bserum\b|\bmoisturiz/i, type: "Skincare" },

        // Sports & Fitness
        { pattern: /\byoga\s*mat\b/i, type: "Yoga Mat" },
        { pattern: /\bdumbbell\b|\bweight/i, type: "Weights" },
        { pattern: /\bfitness\b/i, type: "Fitness Equipment" },
        { pattern: /\bsport/i, type: "Sports Equipment" },

        // Garden
        {
          pattern: /\bgarden\s*(tool|hose|furniture)/i,
          type: "Garden Equipment",
        },
        { pattern: /\bplant\s*(pot|stand)/i, type: "Plant Accessories" },

        // Toys
        { pattern: /\btoy/i, type: "Toy" },
        { pattern: /\bgame\b/i, type: "Game" },
        { pattern: /\bpuzzle\b/i, type: "Puzzle" },
        { pattern: /\blego\b/i, type: "Building Blocks" },

        // Baby
        { pattern: /\bbaby\b/i, type: "Baby Product" },

        // Automotive
        { pattern: /\bcar\s*(seat|mat|cover|charger)/i, type: "Car Accessory" },
      ];

      // Specifics extraction patterns
      this.specificsPatterns = {
        // Common
        Color:
          /\b(black|white|red|blue|green|yellow|orange|pink|purple|brown|grey|gray|gold|silver|beige|navy|sage|cream|teal|burgundy|maroon|olive|coral|turquoise|lavender|mint|rose|charcoal)\b/i,
        Material:
          /\b(cotton|wool|leather|polyester|silk|linen|nylon|denim|fleece|velvet|suede|canvas|mesh|foam|memory\s*foam|metal|wood|wooden|plastic|glass|rubber|silicone|bamboo|ceramic|stainless\s*steel)\b/i,
        Size: /\b(xxs|xs|s|m|l|xl|xxl|xxxl|one\s*size|small|medium|large|extra\s*large|uk\s*\d+|eu\s*\d+|us\s*\d+)\b/i,

        // Dimensions
        Length:
          /(\d+(?:\.\d+)?)\s*(?:cm|mm|m|inch|inches|"|'')?\s*(?:long|length)/i,
        Width:
          /(\d+(?:\.\d+)?)\s*(?:cm|mm|m|inch|inches|"|'')?\s*(?:wide|width)/i,
        Height:
          /(\d+(?:\.\d+)?)\s*(?:cm|mm|m|inch|inches|"|'')?\s*(?:tall|height|high)/i,
        Dimensions:
          /(\d+(?:\.\d+)?\s*x\s*\d+(?:\.\d+)?(?:\s*x\s*\d+(?:\.\d+)?)?)\s*(?:cm|mm|m|in)?/i,

        // Electronics
        "Screen Size":
          /(\d+(?:\.\d+)?)\s*(?:inch|"|''|inches?)\s*(?:screen|display)?/i,
        "Battery Life": /(\d+)\s*(?:hour|hr|h)\s*(?:battery)?/i,
        Connectivity:
          /(wifi|wi-fi|bluetooth|usb|usb-c|hdmi|wireless|5g|4g|lte|nfc)/i,
        Wattage: /(\d+)\s*(?:w|watt)/i,
        Voltage: /(\d+)\s*(?:v|volt)/i,

        // Bedding specific
        "Fill Material":
          /\b(memory\s*foam|down|feather|polyester\s*fill|hollow\s*fibre|microfibre)\b/i,
        Firmness: /\b(soft|medium|firm|extra\s*firm)\b/i,
        "Thread Count": /(\d+)\s*(?:thread\s*count|tc)/i,

        // Clothing specific
        "Sleeve Length":
          /\b(short\s*sleeve|long\s*sleeve|sleeveless|3\/4\s*sleeve)\b/i,
        "Fit Type":
          /\b(slim\s*fit|regular\s*fit|loose\s*fit|skinny|oversized|relaxed|tailored)\b/i,
        Neckline:
          /\b(crew\s*neck|v-neck|round\s*neck|polo|collar|turtle\s*neck|scoop)\b/i,
        Pattern:
          /\b(striped|plaid|floral|solid|printed|graphic|checkered|polka\s*dot|plain)\b/i,

        // Footwear specific
        "UK Shoe Size": /\buk\s*(\d+(?:\.\d+)?)\b/i,
        "EU Shoe Size": /\beu\s*(\d+(?:\.\d+)?)\b/i,
        "Closure Type":
          /\b(lace[\s-]?up|slip[\s-]?on|velcro|zip|buckle|elastic)\b/i,

        // Features
        "Number of Pieces": /(\d+)\s*(?:pack|piece|pcs|set|count)/i,
        "Weight Capacity":
          /(\d+)\s*(?:kg|lb|stone)\s*(?:capacity|weight|load)/i,
        Waterproof: /\b(waterproof|water[\s-]?resistant|splash[\s-]?proof)\b/i,

        // Age/Gender
        "Age Group":
          /\b(adult|kids|children|toddler|baby|infant|teen|senior)\b/i,
        Department: /\b(men|women|unisex|boys|girls|mens|womens)\b/i,

        // Condition
        "Item Condition": /\b(new|brand\s*new|sealed|unopened)\b/i,
      };
    }

    detectType(title) {
      const text = (title || "").toLowerCase();
      for (const { pattern, type } of this.typePatterns) {
        if (pattern.test(text)) return type;
      }
      return "General Product";
    }

    extractAllSpecifics(title, description) {
      const specifics = {};
      const searchText = (
        (title || "") +
        " " +
        (description || "")
      ).toLowerCase();
      const originalText = (title || "") + " " + (description || "");

      // Extract brand
      for (const [brand, pattern] of Object.entries(this.brands)) {
        if (pattern.test(searchText)) {
          specifics.Brand = brand;
          break;
        }
      }

      // Detect and set Type
      specifics.Type = this.detectType(title);

      // Extract all other specifics
      for (const [key, pattern] of Object.entries(this.specificsPatterns)) {
        const match = originalText.match(pattern);
        if (match) {
          let value = (match[1] || match[0]).trim();
          // Capitalize first letter
          value = value.charAt(0).toUpperCase() + value.slice(1).toLowerCase();
          specifics[key] = value;
        }
      }

      return specifics;
    }
  }

  const enhancedSpecificsGenerator = new EnhancedSpecificsGenerator();

  // =============================================================================
  // INTELLIGENT CATEGORY MATCHER - LOOPS THROUGH ALL CATEGORIES
  // =============================================================================

  class IntelligentCategoryMatcher {
    constructor() {
      this.categories = [];
      this.tokenIndex = new Map();
      this.isInitialized = false;

      // Product-to-category mappings
      this.productToCategory = {
        // Bedding / Pillows
        pillow: ["pillows", "bedding"],
        "memory foam pillow": ["pillows", "bedding"],
        "neck pillow": ["pillows", "bedding"],
        "orthopaedic pillow": ["pillows", "bedding"],
        "orthopedic pillow": ["pillows", "bedding"],
        "bed pillow": ["pillows", "bedding"],
        duvet: ["duvets", "bedding"],
        mattress: ["mattresses", "beds mattresses"],
        bedding: ["bedding", "bed linens"],
        blanket: ["blankets throws", "bedding"],
        throw: ["blankets throws", "bedding"],
        cushion: ["cushions", "home decor"],

        // Tablet cases
        "ipad case": [
          "tablet ebook reader accs",
          "cases covers keyboard folios",
        ],
        ipad: ["tablet", "ebook reader"],
        "tablet case": ["tablet ebook reader accs", "cases covers"],

        // Phone cases
        "phone case": ["mobile phone accessories", "cases covers skins"],
        "samsung case": ["mobile phone accessories", "cases covers skins"],
        "iphone case": ["mobile phone accessories", "cases covers skins"],
        "galaxy z fold": ["mobile phone accessories", "cases covers skins"],
        "galaxy z flip": ["mobile phone accessories", "cases covers skins"],

        // Screen protectors
        "screen protector": ["screen protectors"],
        "screen protection": ["screen protectors"],
        "tempered glass": ["screen protectors"],

        // Pet products
        "dog clippers": [
          "pet supplies",
          "dog grooming",
          "clippers scissors shears",
        ],
        "dog grooming": ["pet supplies", "dog grooming"],
        "pet grooming": ["pet supplies", "grooming"],
        "dog trimmer": [
          "pet supplies",
          "dog grooming",
          "clippers scissors shears",
        ],
        "cat grooming": ["pet supplies", "cat grooming"],
        "pet clippers": ["pet supplies", "grooming", "clippers"],

        // Home - Fans
        fan: ["fans", "portable fans"],
        "standing fan": ["portable fans", "pedestal fans"],
        "portable fan": ["portable fans"],
        "desk fan": ["portable fans"],
        "tower fan": ["portable fans"],
        "ceiling fan": ["ceiling fans"],

        // Home - Curtains
        curtains: ["curtains", "pelmets"],
        curtain: ["curtains"],
        blinds: ["blinds"],
        drapes: ["curtains", "drapes"],

        // Car/Vehicle products
        "car mattress": [
          "car accessories",
          "touring travel",
          "inflatable mattresses airbeds",
        ],
        "car travel": ["car accessories", "touring travel"],
        "suv mattress": [
          "car accessories",
          "touring travel",
          "inflatable mattresses airbeds",
        ],
        "air mattress": [
          "inflatable mattresses airbeds",
          "air beds",
          "outdoor sleeping",
        ],
        "inflatable mattress": ["inflatable mattresses airbeds", "air beds"],
        "inflatable air mattress": [
          "inflatable mattresses airbeds",
          "air beds",
        ],

        // Furniture
        "bed risers": [
          "furniture risers",
          "beds mattresses",
          "furniture accessories",
        ],
        "bed riser": ["furniture risers", "beds mattresses"],
        "furniture risers": ["furniture risers", "furniture accessories"],
        "bed lift": ["furniture risers", "beds mattresses"],

        // Electronics
        headphones: ["headphones", "earphones"],
        earbuds: ["headphones", "earphones"],
        speaker: ["speakers", "bluetooth speakers"],
        charger: ["chargers", "chargers sync cables"],
        keyboard: ["keyboards"],
        mouse: ["mice", "mouse"],
        "cd player": ["cd players", "personal cd players", "portable audio"],
        "portable cd": ["cd players", "personal cd players", "portable audio"],
        discman: ["cd players", "personal cd players"],

        // Gaming Controllers
        controller: ["controllers", "gamepads", "video game accessories"],
        gamepad: ["controllers", "gamepads", "video game accessories"],
        "xbox controller": ["xbox accessories", "controllers", "gamepads"],
        xbox: ["xbox accessories", "xbox one", "xbox series"],
        "ps4 controller": ["playstation accessories", "controllers"],
        "ps-4": ["playstation accessories", "ps4"],
        "wired controller": ["controllers", "gamepads"],

        // Urns / Memorial
        urn: ["urns", "cremation urns", "memorial"],
        "cremation urn": ["urns", "cremation urns", "funeral"],
        "human ashes": ["urns", "cremation", "memorial"],
        "pet memorial": ["pet memorial", "urns"],
        "pet urn": ["pet memorial", "urns"],

        // Barn Door Hardware
        "barn door": ["door hardware", "barn door hardware", "sliding door"],
        "sliding door": ["door hardware", "sliding door hardware"],
        "door hardware": ["door hardware", "handles"],
        "door track": ["door hardware", "sliding door"],
        "sliding barn door": ["door hardware", "barn door hardware"],
        "barn door kit": ["door hardware", "barn door hardware"],
        "double barn door": ["door hardware", "barn door hardware"],
        "ceiling mount": ["door hardware", "barn door hardware"],

        // Baby Play Mat / Floor Mat
        "play mat": ["playmats", "baby toys activities"],
        playmat: ["playmats", "baby toys activities"],
        "baby play mat": ["playmats", "baby toys activities"],
        "foam play mat": ["playmats", "baby toys activities"],
        "floor mat": ["playmats", "floor mats"],
        "crawling mat": ["playmats", "baby toys activities"],
        "baby mat": ["playmats", "baby toys activities"],
        "foldable mat": ["playmats", "floor mats"],

        // Caster Wheels
        caster: ["casters", "caster wheels"],
        "caster wheels": ["casters", "caster wheels"],
        casters: ["casters", "caster wheels"],
        "industrial caster": ["casters", "caster wheels"],
        "heavy duty caster": ["casters", "caster wheels"],

        // Decorative Box / Storage Box
        "decorative box": [
          "storage boxes",
          "decorative boxes",
          "trinket boxes",
        ],
        "glass box": ["storage boxes", "decorative boxes", "jewellery boxes"],
        "trinket box": ["trinket boxes", "decorative boxes"],
        "jewellery box": ["jewellery boxes"],
        "lidded box": ["storage boxes", "decorative boxes"],
        "vintage box": ["storage boxes", "decorative boxes"],

        // Drain Rods / Drain Cleaning
        "drain rod": ["drain cleaning", "plumbing tools"],
        "drain rods": ["drain cleaning", "plumbing tools"],
        "drain cleaning": ["drain cleaning", "pipe tools"],
        "pressure washer": ["pressure washers", "cleaning"],

        // Footmuff / Cosytoes
        footmuff: ["cosytoes", "footmuffs", "pushchair accessories"],
        cosytoes: ["cosytoes", "footmuffs", "pushchair accessories"],
        cosytoe: ["cosytoes", "footmuffs"],
        "baby footmuff": ["cosytoes", "footmuffs", "pushchair accessories"],
        handmuff: ["pushchair accessories"],

        // Galaxy Z Fold Case
        "galaxy z fold": ["mobile phone accessories", "cases covers skins"],
        "z fold case": ["mobile phone accessories", "cases covers skins"],
        "samsung galaxy": ["mobile phone accessories", "cases covers skins"],
        "galaxy fold": ["mobile phone accessories", "cases covers skins"],

        // Plumbing
        cistern: ["cisterns", "toilet parts", "toilets bidets"],
        "flush cistern": ["cisterns", "toilet parts"],
        toilet: ["toilets", "toilets bidets"],

        // Garden Covers
        "table cover": ["furniture covers", "garden furniture covers"],
        "garden cover": ["furniture covers", "outdoor furniture covers"],
        "patio cover": ["furniture covers", "outdoor furniture covers"],
        "hot tub cover": ["hot tub covers", "spa covers"],
        "hot tub": ["hot tubs", "spa", "swimming pools hot tubs"],

        // Garment Bags / Luggage
        "suit bag": ["garment bags", "suit carriers"],
        "garment bag": ["garment bags", "suit carriers"],

        // Pet Training
        "bark collar": ["bark control", "dog training"],
        "bark control": ["bark control", "dog training"],
        "dog collar": ["dog collars", "pet supplies"],
        "dog steps": ["ramps stairs", "pet supplies"],
        "pet stairs": ["ramps stairs", "pet supplies"],

        // Humidity / Climate
        humidifier: ["humidifiers", "indoor air quality"],
        "humidity controller": ["thermostats", "climate control"],
        thermostat: ["thermostats", "hvac"],

        // Party Supplies
        "plastic cups": ["party tableware", "disposable tableware"],
        "party cups": ["party tableware", "party supplies"],

        // Clothing
        shorts: ["shorts", "mens shorts"],
        "linen shorts": ["shorts", "mens clothing"],
        "mens shorts": ["shorts", "mens clothing"],
        coat: ["coats jackets", "mens coats", "outerwear"],
        "trench coat": ["coats jackets", "trench coats"],
        overcoat: ["coats jackets", "overcoats"],
        "wool coat": ["coats jackets", "wool coats"],
        jacket: ["jackets", "coats jackets"],
        "mens coat": ["coats jackets", "mens coats"],
        "womens coat": ["coats jackets", "womens coats"],

        // Office Equipment
        shredder: ["shredders", "paper shredders", "office equipment"],
        "paper shredder": ["shredders", "paper shredders"],
        "cross cut shredder": ["shredders", "paper shredders"],

        // Headsets / Audio
        headset: ["headsets", "headphones", "pc headsets"],
        "pc headset": ["headsets", "pc headsets"],
        "business headset": ["headsets", "office headsets"],
        "bluetooth headset": ["headphones", "bluetooth headsets"],
        "wireless headset": ["headphones", "wireless headsets"],

        // Hand Dryer
        "hand dryer": ["hand dryers", "bathroom accessories", "commercial"],
        "commercial hand dryer": ["hand dryers", "bathroom"],

        // Thermostat / Heating
        thermostat: ["thermostats", "heating cooling"],
        "wifi thermostat": ["thermostats", "smart thermostats"],
        "underfloor heating": ["underfloor heating", "heating systems"],
        "floor heating": ["underfloor heating", "heating systems"],

        // Hook-on High Chair
        "hook on high chair": ["high chairs", "baby feeding"],
        "hook on chair": ["high chairs", "baby feeding"],
        "hook-on": ["high chairs", "baby feeding"],
        "portable high chair": ["high chairs", "baby feeding"],

        // Cat Repellent / Pest Control
        "cat repellent": ["pest repellers", "pest control", "ultrasonic"],
        "cat scarer": ["pest repellers", "pest control"],
        "cat deterrent": ["pest repellers", "pest control"],
        "animal repellent": ["pest repellers", "pest control"],
        "ultrasonic repeller": ["pest repellers", "ultrasonic"],
        "pest repeller": ["pest repellers", "pest control"],

        // TV Remote
        "tv remote": ["remote controls", "tv accessories"],
        "smart tv remote": ["remote controls", "tv accessories"],
        "lg remote": ["remote controls", "tv accessories"],
        "replacement remote": ["remote controls", "remotes"],

        // Passport / Travel Accessories
        "passport cover": ["passport holders", "travel accessories"],
        "passport holder": ["passport holders", "travel accessories"],
        "passport case": ["passport holders", "travel accessories"],

        // Laptop Bags
        "laptop sleeve": ["laptop bags cases", "laptop sleeves"],
        "laptop case": ["laptop bags cases", "laptop cases"],
        "laptop bag": ["laptop bags cases"],

        // Watches
        watch: ["watches", "wristwatches"],
        wristwatch: ["watches", "wristwatches"],
        "mens watch": ["mens watches", "wristwatches"],
        "womens watch": ["womens watches", "wristwatches"],
        "women's watch": ["womens watches", "wristwatches"],
        "men's watch": ["mens watches", "wristwatches"],

        // Room Divider
        "room divider": ["room dividers", "screens"],
        partition: ["room dividers", "screens"],

        // Phone Cases
        "iphone case": ["cases covers skins", "mobile phone accessories"],
        "phone case": ["cases covers skins", "mobile phone accessories"],
        "wallet case": ["cases covers skins", "mobile phone accessories"],

        // Appliance Parts
        "dishwasher basket": ["dishwasher parts", "appliance parts"],
        "cutlery basket": ["dishwasher parts", "appliance parts"],
        "oven glass": ["cooker parts", "oven parts"],
        "oven lens": ["cooker parts", "oven parts"],
        "lamp cover": ["cooker parts", "lamp covers"],
        "dryer vent": ["tumble dryer parts", "washing machine dryer parts"],
        "vent hose": ["tumble dryer parts", "vent hoses"],

        // Kitchen Appliances
        "pasta maker": ["pasta makers", "small kitchen appliances"],
        "pasta machine": ["pasta makers", "small kitchen appliances"],
        dehumidifier: ["dehumidifiers", "indoor air quality"],

        // Rotary Tool
        "rotary tool": ["rotary tools", "power tools"],
        "multi tool": ["multi tools", "power tools"],

        // Beard Trimmer
        "beard trimmer": ["clippers trimmers", "shavers trimmers"],
        "stubble trimmer": ["clippers trimmers", "beard trimmers"],

        // Remote Control
        "remote control": ["remote controls", "remotes"],
        "replacement remote": ["remote controls", "remotes"],

        // Sauna
        sauna: ["saunas", "steam saunas"],
        "sauna tent": ["saunas", "portable saunas"],

        // Photography
        backdrop: ["backdrops", "photography backgrounds"],
        "photography backdrop": ["backdrops", "backgrounds"],

        // Cassette Player
        "cassette player": ["cassette players", "portable audio"],
        cassette: ["cassette players", "portable audio"],

        // Baby
        "baby sleeping bag": ["sleeping bags sleepsacks", "baby bedding"],
        "sleep sack": ["sleeping bags sleepsacks", "baby bedding"],

        // Water Bottle
        "water bottle": ["water bottles", "sports bottles"],

        // Bed Alarm / Medical
        "bed alarm": ["patient alarms", "medical monitors"],
        "caregiver pager": ["patient alarms", "medical monitors"],

        // Filling Valve / Plumbing
        "filling valve": ["toilet parts", "cistern parts"],
        "cistern valve": ["toilet parts", "cistern parts"],
        geberit: ["toilet parts", "cistern parts"],

        // Outdoor Lights
        "outdoor lights": ["outdoor lighting", "wall lights"],
        "wall lights": ["wall lights", "outdoor lighting"],
        "motion sensor light": ["security lights", "outdoor lighting"],

        // Blackout Curtains
        "blackout curtains": ["curtains", "blackout curtains"],
        "kids curtains": ["curtains", "childrens curtains"],

        // Medical / Mobility
        crutch: ["crutches", "mobility", "walking equipment"],
        "knee crutch": ["crutches", "mobility", "walking equipment"],
        "mobility aid": ["mobility", "walking equipment", "medical mobility"],
        "walking aid": ["mobility", "walking equipment"],
        wheelchair: ["wheelchairs", "mobility"],
        walker: ["walkers", "mobility", "walking frames"],

        // Baby Products
        playpen: ["playpens", "baby", "nursery"],
        "baby playpen": ["playpens", "baby", "nursery"],
        "high chair": ["high chairs", "baby feeding", "highchairs"],
        highchair: ["high chairs", "baby feeding"],
        stokke: ["high chairs", "baby feeding"],
        "baby gate": ["baby safety", "safety gates"],
        "baby monitor": ["baby monitors", "baby safety"],
        pram: ["prams", "pushchairs", "strollers"],
        pushchair: ["pushchairs", "prams", "strollers"],
        "car seat": ["car seats", "baby travel"],

        // Kitchen Tools
        grater: ["graters", "food preparation", "kitchen tools"],
        "cheese grater": ["graters", "food preparation"],
        slicer: ["slicers", "mandolines", "food preparation"],
        "vegetable slicer": ["slicers", "mandolines", "food preparation"],
        peeler: ["peelers", "food preparation"],
        chopper: ["choppers", "food preparation"],
        blender: ["blenders", "small kitchen appliances"],
        mixer: ["mixers", "small kitchen appliances"],

        // Air Quality / Appliances
        "air purifier": ["air purifiers", "indoor air quality"],
        purifier: ["air purifiers", "purifiers"],
        "hepa filter": ["air purifiers", "filters"],
        "replacement filter": ["filters", "replacement parts"],
        humidifier: ["humidifiers", "indoor air quality"],
        dehumidifier: ["dehumidifiers", "indoor air quality"],

        // Door/Thermal Curtains
        "door curtain": ["curtains", "door curtains"],
        "thermal curtain": ["curtains", "thermal curtains"],
        "blackout curtain": ["curtains", "blackout curtains"],
      };

      // Synonyms for token expansion
      this.synonyms = {
        case: ["cases", "cover", "covers", "folios"],
        curtains: ["curtain"],
        curtain: ["curtains"],
        fan: ["fans"],
        clippers: ["clipper", "trimmer", "trimmers", "scissors", "shears"],
        trimmer: ["clipper", "clippers", "trimmers"],
        grooming: ["groom"],
        dog: ["dogs", "pet", "canine"],
        mattress: ["mattresses", "airbed", "airbeds"],
        inflatable: ["airbeds", "airbed"],
        car: ["vehicle", "suv"],
        travel: ["touring", "travelling"],
        risers: ["riser"],
        bed: ["beds"],
        protector: ["protectors", "protection"],
        portable: ["travel", "compact"],
        standing: ["pedestal", "stand"],
        ipad: ["tablet", "tablets"],
        pillow: ["pillows"],
        ortopedic: ["orthopaedic", "orthopedic"],
        orthopedic: ["orthopaedic"],
        duvet: ["duvets"],
        cushion: ["cushions"],
        blanket: ["blankets"],
        // New synonyms
        crutch: ["crutches"],
        playpen: ["playpens", "play pen", "play yard"],
        highchair: ["high chair", "highchairs"],
        grater: ["graters"],
        slicer: ["slicers"],
        filter: ["filters"],
        purifier: ["purifiers"],
        baby: ["infant", "toddler"],
        chair: ["chairs"],
      };

      // Domain keywords for bonus scoring
      this.domainKeywords = {
        pet: ["pet", "dog", "cat", "animal"],
        mobile: ["mobile", "phone", "communication"],
        tablet: ["tablet", "ebook", "reader", "computers"],
        home: ["home", "furniture", "diy", "garden", "bedding", "curtain"],
        car: ["vehicle", "car", "automotive", "motor"],
        camping: ["camping", "outdoor", "hiking", "sporting"],
        beauty: ["health", "beauty"],
        baby: ["baby", "nursery", "infant", "toddler"],
        medical: ["medical", "mobility", "health"],
        kitchen: ["kitchen", "cookware", "food preparation"],
        appliance: ["appliance", "heating", "cooling", "air"],
      };
    }

    tokenize(text) {
      if (!text) return [];
      return text
        .toString()
        .toLowerCase()
        .replace(/rrp\s*£\d+/gi, "")
        .replace(/\d+x\d+/gi, "")
        .replace(/\d+[-]?inch/gi, "")
        .replace(/[^\w\s-]/g, " ")
        .replace(/\s+/g, " ")
        .trim()
        .split(/\s+/)
        .filter((word) => word.length > 1 && !STOPWORDS.has(word));
    }

    expandTokens(text, tokens) {
      const expanded = new Set(tokens);
      const textLower = text.toLowerCase();

      for (const [productTerm, categoryTerms] of Object.entries(
        this.productToCategory
      )) {
        if (textLower.includes(productTerm)) {
          for (const catTerm of categoryTerms) {
            this.tokenize(catTerm).forEach((t) => expanded.add(t));
          }
        }
      }

      for (const token of [...expanded]) {
        if (this.synonyms[token]) {
          this.synonyms[token].forEach((syn) => expanded.add(syn));
        }
      }

      // Remove standalone "air" if we have "mattress"
      if (textLower.includes("mattress") && expanded.has("air")) {
        expanded.delete("air");
      }

      return Array.from(expanded);
    }

    build(categoryMapData) {
      if (!categoryMapData || categoryMapData.length === 0) return;

      this.tokenIndex.clear();
      this.categories = [];

      categoryMapData.forEach((cat, idx) => {
        const categoryId =
          cat["Category ID"] || cat.CategoryID || cat.categoryId;
        const categoryPath =
          cat["Category Path"] || cat.CategoryPath || cat.categoryPath || "";

        if (!categoryId || !categoryPath) return;

        const pathStr = categoryPath.toString();
        const pathParts = pathStr.split(">").map((p) => p.trim());
        const depth = pathParts.length;
        const leafCategory = pathParts[pathParts.length - 1];

        const pathTokens = this.tokenize(pathStr);
        const leafTokens = this.tokenize(leafCategory);

        const categoryData = {
          id: categoryId.toString(),
          path: pathStr,
          tokens: pathTokens,
          leafTokens: leafTokens,
          depth: depth,
          leaf: leafCategory.toLowerCase(),
          index: idx,
        };

        this.categories.push(categoryData);

        // Index ALL tokens
        pathTokens.forEach((token) => {
          if (!this.tokenIndex.has(token)) {
            this.tokenIndex.set(token, new Set());
          }
          this.tokenIndex.get(token).add(idx);
        });
      });

      this.isInitialized = true;
      console.log(
        `✓ Category matcher initialized with ${this.categories.length} categories`
      );
    }

    detectDomain(title) {
      const titleLower = title.toLowerCase();

      if (/\b(dog|cat|pet|puppy|kitten)\b/.test(titleLower)) return "pet";
      if (/\b(ipad|tablet)\b/.test(titleLower)) return "tablet";
      if (/\b(phone|iphone|samsung|galaxy|mobile)\b/.test(titleLower))
        return "mobile";
      if (/\b(car|suv|vehicle|automotive)\b/.test(titleLower)) return "car";
      if (/\b(camping|outdoor|hiking)\b/.test(titleLower)) return "camping";
      if (
        /\b(curtain|furniture|bed|sofa|chair|fan|home|pillow|duvet|bedding)\b/.test(
          titleLower
        )
      )
        return "home";
      if (/\b(makeup|skincare|beauty|cosmetic)\b/.test(titleLower))
        return "beauty";

      return null;
    }

    match(
      title,
      description = "",
      excelCategory = "",
      excelSubCategory = "",
      rowIndex = -1
    ) {
      if (!this.isInitialized || this.categories.length === 0) {
        return {
          categoryId: "47155",
          categoryPath: "Default",
          score: 0,
          confidence: "none",
        };
      }

      if (!title) {
        return {
          categoryId: "47155",
          categoryPath: "Default",
          score: 0,
          confidence: "none",
        };
      }

      const titleLower = title.toLowerCase();
      const titleTokens = this.tokenize(title);
      const expandedTokens = this.expandTokens(title, titleTokens);
      const detectedDomain = this.detectDomain(title);

      // Get Google Search result if available
      let googleSearchResult = null;
      if (this.googleSearchResults && rowIndex >= 0) {
        googleSearchResult = this.googleSearchResults.get(rowIndex);
        if (googleSearchResult && googleSearchResult.suggestedCategory) {
          console.log(
            `Google suggests category: ${googleSearchResult.suggestedCategory} for row ${rowIndex}`
          );
        }
      }

      // Tokenize Excel category and subcategory for matching
      const excelCategoryLower = excelCategory.toLowerCase();
      const excelSubCategoryLower = excelSubCategory.toLowerCase();
      const excelCategoryTokens = this.tokenize(
        excelCategory + " " + excelSubCategory
      );

      if (excelCategory || excelSubCategory) {
        console.log(
          `Category Matching: Excel Category="${excelCategory}" SubCategory="${excelSubCategory}"`
        );
      }

      if (titleTokens.length === 0) {
        return {
          categoryId: "47155",
          categoryPath: "Default",
          score: 0,
          confidence: "none",
        };
      }

      // Score ALL matching categories - NO EARLY STOPPING
      const scores = new Map();

      // First, add tokens from Excel category/subcategory to expanded tokens
      const allTokens = [
        ...new Set([...expandedTokens, ...excelCategoryTokens]),
      ];

      for (const token of allTokens) {
        if (!this.tokenIndex.has(token)) continue;

        const isOriginal = titleTokens.includes(token);
        const isFromExcelCategory = excelCategoryTokens.includes(token);

        // Loop through ALL categories that have this token
        for (const catIdx of this.tokenIndex.get(token)) {
          const cat = this.categories[catIdx];

          if (!scores.has(catIdx)) {
            scores.set(catIdx, {
              leaf: 0,
              path: 0,
              matched: new Set(),
              excelBonus: 0,
            });
          }

          const scoreData = scores.get(catIdx);

          if (cat.leafTokens.includes(token)) {
            scoreData.leaf += isOriginal ? 50 : isFromExcelCategory ? 60 : 35;
            scoreData.matched.add(token + "*");
          } else if (cat.tokens.includes(token)) {
            scoreData.path += isOriginal ? 10 : isFromExcelCategory ? 15 : 5;
            scoreData.matched.add(token);
          }
        }
      }

      // Calculate final scores for ALL matched categories
      const results = [];

      scores.forEach((scoreData, catIdx) => {
        const cat = this.categories[catIdx];
        const pathLower = cat.path.toLowerCase();
        const leafLower = cat.leaf.toLowerCase();

        let leafScore = scoreData.leaf * 2;
        let pathScore = scoreData.path;
        let depthBonus = cat.depth * 10;

        // Domain bonus
        let domainBonus = 0;
        if (detectedDomain && this.domainKeywords[detectedDomain]) {
          if (
            this.domainKeywords[detectedDomain].some((kw) =>
              pathLower.includes(kw)
            )
          ) {
            domainBonus = 50;
          }
        }

        // === PRODUCT LOOKUP BONUS - DISABLED ===
        // API-based lookup removed, using title-based matching only
        let googleBonus = 0;

        // === EXCEL CATEGORY MATCHING BONUS - HELPER (NOT DOMINANT) ===
        let excelCategoryBonus = 0;

        // Map Excel categories to eBay path keywords
        const categoryMappings = {
          "home & kitchen": ["home", "furniture", "kitchen", "appliances"],
          "tools & diy": ["diy", "tools", "hardware", "building"],
          baby: ["baby", "nursery", "feeding"],
          electronics: [
            "electronics",
            "mobile",
            "phones",
            "cameras",
            "sound",
            "vision",
          ],
          "sports & leisure": ["sporting", "fitness", "leisure"],
          garden: ["garden", "patio", "outdoor"],
          clothing: ["clothes", "shoes", "accessories"],
          "health & beauty": ["health", "beauty"],
          automotive: ["vehicle", "car", "automotive", "motor"],
        };

        const subCategoryMappings = {
          "bedding & bath": ["bedding", "bath", "towel", "linen"],
          hardware: ["hardware", "door", "handle", "hinge"],
          "building materials": ["building", "hardware", "construction"],
          "baby gear": ["baby", "play", "mat", "bouncer"],
          "home accessories": ["home", "decor", "storage", "box"],
          dogs: ["dog", "pet", "grooming", "clipper"],
          "gardening tools": ["garden", "tools", "drain", "hose"],
          "mobile phones & accessories": ["mobile", "phone", "case", "cover"],
          "strollers & car seats": [
            "pushchair",
            "stroller",
            "footmuff",
            "car seat",
          ],
          "baby feeding": ["feeding", "high chair", "bottle"],
          "tyres & rims": ["tyre", "tire", "wheel", "rim", "vehicle"],
          "car parts": ["vehicle", "car", "automotive", "parts"],
        };

        // Check if Excel category matches eBay path - moderate bonus
        if (excelCategoryLower) {
          for (const [excelCat, pathKeywords] of Object.entries(
            categoryMappings
          )) {
            if (
              excelCategoryLower.includes(excelCat) ||
              excelCat.includes(excelCategoryLower)
            ) {
              if (pathKeywords.some((kw) => pathLower.includes(kw))) {
                excelCategoryBonus += 80; // Moderate bonus for category match
              }
            }
          }
        }

        // Check if Excel subcategory matches eBay path - moderate bonus
        if (excelSubCategoryLower) {
          for (const [excelSubCat, pathKeywords] of Object.entries(
            subCategoryMappings
          )) {
            if (
              excelSubCategoryLower.includes(excelSubCat) ||
              excelSubCat.includes(excelSubCategoryLower)
            ) {
              if (
                pathKeywords.some(
                  (kw) => pathLower.includes(kw) || leafLower.includes(kw)
                )
              ) {
                excelCategoryBonus += 100; // Moderate bonus for subcategory match
              }
            }
          }
        }

        // Direct keyword matching from Excel category/subcategory
        for (const token of excelCategoryTokens) {
          if (token.length > 2) {
            if (leafLower.includes(token)) excelCategoryBonus += 40;
            else if (pathLower.includes(token)) excelCategoryBonus += 20;
          }
        }

        // Product-specific bonuses
        let productBonus = 0;

        // iPad/Tablet case
        if (/\bipad\b/.test(titleLower) || /\btablet\b/.test(titleLower)) {
          if (pathLower.includes("tablet") || pathLower.includes("ebook"))
            productBonus += 50;
          if (
            leafLower.includes("cases") ||
            leafLower.includes("covers") ||
            leafLower.includes("folios")
          )
            productBonus += 40;
        }

        // Phone case
        if (
          /\bphone\b/.test(titleLower) ||
          /\bgalaxy\b/.test(titleLower) ||
          /\biphone\b/.test(titleLower)
        ) {
          if (pathLower.includes("mobile") || pathLower.includes("phone"))
            productBonus += 50;
          if (
            leafLower.includes("cases") ||
            leafLower.includes("covers") ||
            leafLower.includes("skins")
          )
            productBonus += 40;
        }

        // Dog/Pet grooming
        if (/\bdog\b/.test(titleLower) && pathLower.includes("dog"))
          productBonus += 40;
        if (/\bgrooming\b/.test(titleLower) && pathLower.includes("grooming"))
          productBonus += 50;
        if (
          (/\bclippers?\b/.test(titleLower) ||
            /\btrimmer\b/.test(titleLower)) &&
          (leafLower.includes("clipper") ||
            leafLower.includes("trimmer") ||
            leafLower.includes("scissors") ||
            leafLower.includes("shears"))
        ) {
          productBonus += 40;
        }

        // Fan
        if (/\bfan\b/.test(titleLower)) {
          if (leafLower.includes("fans") && pathLower.includes("heating"))
            productBonus += 60;
          if (leafLower.includes("portable")) productBonus += 40;
        }

        // Curtains
        if (/\bcurtains?\b/.test(titleLower)) {
          if (leafLower.includes("curtains") && leafLower.includes("pelmets"))
            productBonus += 50;
          if (pathLower.includes("curtains") && pathLower.includes("blinds"))
            productBonus += 30;
        }

        // Air mattress / Inflatable
        if (
          /\bmattress\b/.test(titleLower) ||
          /\binflatable\b/.test(titleLower)
        ) {
          if (
            leafLower.includes("inflatable") &&
            (leafLower.includes("mattress") || leafLower.includes("airbed"))
          ) {
            productBonus += 100;
          } else if (
            leafLower.includes("airbed") ||
            leafLower.includes("air bed")
          ) {
            productBonus += 80;
          }
        }

        // Bed risers
        if (/\briser/i.test(titleLower)) {
          if (leafLower.includes("riser")) productBonus += 150;
          else if (
            leafLower.includes("accessories") &&
            (pathLower.includes("bed") || pathLower.includes("furniture"))
          ) {
            productBonus += 80;
          }
        }

        // Pillows / Bedding
        if (/\bpillow\b/.test(titleLower)) {
          if (leafLower === "pillows") {
            productBonus += 150;
          } else if (
            leafLower.includes("pillows") &&
            !leafLower.includes("case") &&
            !leafLower.includes("ring")
          ) {
            productBonus += 100;
          }
          if (
            leafLower.includes("pillow case") ||
            leafLower.includes("pillowcase")
          ) {
            productBonus -= 80;
          }
          if (pathLower.includes("bedding")) {
            productBonus += 40;
          }
          if (
            (pathLower.includes("children") ||
              pathLower.includes("baby") ||
              pathLower.includes("nursery")) &&
            !/\b(child|children|kids|baby|nursery|toddler)\b/.test(titleLower)
          ) {
            productBonus -= 30;
          }
        }

        // Duvet / Bedding
        if (/\bduvet\b/.test(titleLower)) {
          if (leafLower.includes("duvet")) productBonus += 100;
          if (pathLower.includes("bedding")) productBonus += 40;
        }

        // Cushion
        if (/\bcushion\b/.test(titleLower)) {
          if (leafLower.includes("cushion")) productBonus += 100;
        }

        // Travel (car)
        if (
          /\btravel\b/.test(titleLower) &&
          (leafLower.includes("travel") || leafLower.includes("touring"))
        ) {
          productBonus += 40;
        }

        // === NEW PRODUCT BONUSES ===

        // Crutch / Mobility Aid
        if (/\bcrutch\b/.test(titleLower)) {
          if (leafLower.includes("crutch")) productBonus += 150;
          if (pathLower.includes("mobility") || pathLower.includes("walking"))
            productBonus += 50;
        }
        if (/\bmobility\s*aid\b/.test(titleLower)) {
          if (pathLower.includes("mobility")) productBonus += 80;
        }

        // Playpen
        if (/\bplaypen\b/.test(titleLower)) {
          if (leafLower.includes("playpen") || leafLower.includes("play yard"))
            productBonus += 150;
          if (pathLower.includes("baby") || pathLower.includes("nursery"))
            productBonus += 50;
        }

        // High Chair
        if (
          /\bhigh\s*chair\b/.test(titleLower) ||
          /\bhighchair\b/.test(titleLower)
        ) {
          if (
            leafLower.includes("high chair") ||
            leafLower.includes("highchair")
          )
            productBonus += 150;
          if (pathLower.includes("baby") || pathLower.includes("feeding"))
            productBonus += 50;
        }

        // Grater / Slicer
        if (/\bgrater\b/.test(titleLower)) {
          if (leafLower.includes("grater")) productBonus += 150;
          if (
            pathLower.includes("kitchen") ||
            pathLower.includes("food preparation")
          )
            productBonus += 50;
        }
        if (/\bslicer\b/.test(titleLower)) {
          if (leafLower.includes("slicer") || leafLower.includes("mandoline"))
            productBonus += 100;
          if (
            pathLower.includes("kitchen") ||
            pathLower.includes("food preparation")
          )
            productBonus += 40;
        }

        // Air Purifier / Filter
        if (
          /\bair\s*purifier\b/.test(titleLower) ||
          /\bpurifier\b/.test(titleLower)
        ) {
          if (
            leafLower.includes("air purifier") ||
            leafLower.includes("purifiers")
          )
            productBonus += 150;
          if (
            pathLower.includes("air quality") ||
            pathLower.includes("heating")
          )
            productBonus += 50;
        }
        if (/\bfilter\b/.test(titleLower) && /\bair\b/.test(titleLower)) {
          if (
            leafLower.includes("air purifier") ||
            pathLower.includes("air quality")
          )
            productBonus += 80;
          // Penalize car air filters
          if (pathLower.includes("vehicle") || pathLower.includes("car parts"))
            productBonus -= 60;
        }

        // Door Curtain / Thermal Curtain
        if (/\bdoor\s*curtain\b/.test(titleLower)) {
          if (leafLower.includes("curtain")) productBonus += 100;
          if (pathLower.includes("curtains") && pathLower.includes("blinds"))
            productBonus += 50;
        }
        if (/\bthermal\b/.test(titleLower) && /\bcurtain\b/.test(titleLower)) {
          if (leafLower.includes("curtain")) productBonus += 80;
        }

        // === MORE PRODUCT BONUSES ===

        // Gaming Controllers (Xbox, PlayStation)
        if (
          /\bcontroller\b/.test(titleLower) ||
          /\bgamepad\b/.test(titleLower)
        ) {
          // Prefer Video Games controllers over industrial/vehicle controllers
          if (
            pathLower.includes("video game") ||
            pathLower.includes("xbox") ||
            pathLower.includes("playstation")
          ) {
            productBonus += 150;
          }
          if (
            leafLower.includes("controller") &&
            pathLower.includes("video game")
          ) {
            productBonus += 80;
          }
          // Penalize wrong controller categories
          if (
            pathLower.includes("vehicle") ||
            pathLower.includes("power") ||
            pathLower.includes("industrial")
          ) {
            productBonus -= 150;
          }
        }
        if (/\bxbox\b/.test(titleLower)) {
          if (pathLower.includes("xbox") || pathLower.includes("video game"))
            productBonus += 100;
        }
        if (
          /\bps[\-]?4\b/.test(titleLower) ||
          /\bplaystation\b/.test(titleLower)
        ) {
          if (
            pathLower.includes("playstation") ||
            pathLower.includes("video game")
          )
            productBonus += 100;
        }

        // CD Player / Discman
        if (
          /\bcd\s*player\b/.test(titleLower) ||
          /\bdiscman\b/.test(titleLower)
        ) {
          if (
            leafLower.includes("cd player") ||
            leafLower.includes("personal cd")
          ) {
            productBonus += 200; // Very strong for exact match
          }
          if (pathLower.includes("portable audio")) productBonus += 80;
          // Penalize wrong matches
          if (
            pathLower.includes("memorial") ||
            pathLower.includes("funeral") ||
            pathLower.includes("cremation")
          ) {
            productBonus -= 200;
          }
        }

        // Urns / Memorial
        if (/\burn\b/.test(titleLower) || /\burns\b/.test(titleLower)) {
          if (leafLower.includes("urn") || leafLower.includes("cremation")) {
            productBonus += 200;
          }
          if (pathLower.includes("memorial") || pathLower.includes("funeral"))
            productBonus += 80;
        }
        if (/\bcremation\b/.test(titleLower)) {
          if (pathLower.includes("memorial") || pathLower.includes("funeral"))
            productBonus += 100;
        }

        // Barn Door Hardware
        if (/\bbarn\s*door\b/.test(titleLower)) {
          if (leafLower.includes("door") && leafLower.includes("hardware"))
            productBonus += 150;
          if (leafLower.includes("sliding")) productBonus += 50;
          // Penalize vehicle doors
          if (pathLower.includes("vehicle") || pathLower.includes("car"))
            productBonus -= 100;
        }

        // Cistern / Toilet
        if (/\bcistern\b/.test(titleLower)) {
          if (
            leafLower.includes("cistern") ||
            (leafLower.includes("toilet") && leafLower.includes("part"))
          ) {
            productBonus += 150;
          }
          if (
            pathLower.includes("toilet") ||
            pathLower.includes("bathroom") ||
            pathLower.includes("plumbing")
          ) {
            productBonus += 50;
          }
        }

        // Garment Bag / Suit Bag
        if (
          /\bsuit\s*bag\b/.test(titleLower) ||
          /\bgarment\s*bag\b/.test(titleLower)
        ) {
          if (
            leafLower.includes("garment") ||
            leafLower.includes("suit carrier")
          )
            productBonus += 150;
          if (pathLower.includes("luggage")) productBonus += 50;
        }

        // Garden/Patio Table Covers
        if (
          /\btable\s*cover\b/.test(titleLower) ||
          (/\bgarden\b/.test(titleLower) && /\bcover\b/.test(titleLower))
        ) {
          if (
            leafLower.includes("cover") &&
            (pathLower.includes("furniture") || pathLower.includes("garden"))
          ) {
            productBonus += 100;
          }
          if (pathLower.includes("outdoor") || pathLower.includes("patio"))
            productBonus += 50;
        }

        // Hot Tub Cover
        if (/\bhot\s*tub\b/.test(titleLower)) {
          if (leafLower.includes("hot tub")) productBonus += 150;
          if (pathLower.includes("swimming") || pathLower.includes("spa"))
            productBonus += 50;
        }

        // Humidifier
        if (/\bhumidifier\b/.test(titleLower)) {
          if (leafLower.includes("humidifier")) productBonus += 150;
          if (pathLower.includes("air quality")) productBonus += 50;
        }

        // Bark Control / Dog Training
        if (
          /\bbark\b/.test(titleLower) &&
          (/\bcollar\b/.test(titleLower) || /\bcontrol\b/.test(titleLower))
        ) {
          if (leafLower.includes("bark") || leafLower.includes("training"))
            productBonus += 150;
          if (pathLower.includes("dog") || pathLower.includes("pet"))
            productBonus += 50;
        }

        // Dog Steps / Pet Stairs
        if (
          /\bdog\s*steps\b/.test(titleLower) ||
          /\bpet\s*stairs\b/.test(titleLower)
        ) {
          if (
            leafLower.includes("stairs") ||
            leafLower.includes("steps") ||
            leafLower.includes("ramps")
          ) {
            productBonus += 150;
          }
          if (pathLower.includes("pet") || pathLower.includes("dog"))
            productBonus += 60;
        }

        // Party Supplies / Plastic Cups
        if (
          /\bplastic\s*cups\b/.test(titleLower) ||
          /\bparty\s*cups\b/.test(titleLower)
        ) {
          if (leafLower.includes("tableware") || leafLower.includes("cups"))
            productBonus += 100;
          if (pathLower.includes("party")) productBonus += 50;
        }

        // Shorts (Clothing)
        if (/\bshorts\b/.test(titleLower)) {
          if (leafLower.includes("shorts")) productBonus += 100;
          if (/\bmen\b/.test(titleLower) && pathLower.includes("men"))
            productBonus += 50;
          // Prefer general clothing over sports-specific
          if (!/\bgolf\b/.test(titleLower) && pathLower.includes("golf"))
            productBonus -= 30;
        }

        // Wheelchair accessories
        if (/\bwheelchair\b/.test(titleLower)) {
          if (leafLower.includes("wheelchair")) productBonus += 150;
          if (pathLower.includes("mobility")) productBonus += 50;
        }

        // === ADDITIONAL PRODUCT BONUSES ===

        // Coats / Jackets / Outerwear
        if (/\bcoat\b/.test(titleLower) || /\bovercoat\b/.test(titleLower)) {
          if (leafLower.includes("coat") || leafLower.includes("jacket"))
            productBonus += 150;
          if (/\bmen\b/.test(titleLower) && pathLower.includes("men"))
            productBonus += 50;
          if (/\bwomen\b/.test(titleLower) && pathLower.includes("women"))
            productBonus += 50;
          // Penalize sports-specific when not sports
          if (!/\bgolf\b/.test(titleLower) && pathLower.includes("golf"))
            productBonus -= 80;
        }

        // Shredder (Paper Shredder)
        if (/\bshredder\b/.test(titleLower)) {
          if (leafLower.includes("shredder") && pathLower.includes("office"))
            productBonus += 200;
          if (leafLower.includes("paper shredder")) productBonus += 150;
          // Penalize wrong shredder categories
          if (
            pathLower.includes("pump") ||
            pathLower.includes("garden") ||
            pathLower.includes("industrial")
          ) {
            productBonus -= 150;
          }
        }

        // PC Headset
        if (/\bheadset\b/.test(titleLower)) {
          if (leafLower.includes("headset") || leafLower.includes("headphone"))
            productBonus += 150;
          if (/\bpc\b/.test(titleLower) && pathLower.includes("computer"))
            productBonus += 50;
          // Penalize VR headsets when not VR
          if (
            !/\bvr\b/.test(titleLower) &&
            pathLower.includes("virtual reality")
          )
            productBonus -= 100;
        }

        // iPhone / Phone Case
        if (/\biphone\b/.test(titleLower) && /\bcase\b/.test(titleLower)) {
          if (leafLower.includes("cases") || leafLower.includes("covers"))
            productBonus += 200;
          if (pathLower.includes("mobile")) productBonus += 80;
          // Penalize wrong categories
          if (pathLower.includes("group") || pathLower.includes("collectables"))
            productBonus -= 150;
        }

        // Oven / Cooker Parts
        if (
          /\boven\b/.test(titleLower) &&
          (/\bglass\b/.test(titleLower) ||
            /\blens\b/.test(titleLower) ||
            /\bcover\b/.test(titleLower))
        ) {
          if (pathLower.includes("cooker") || pathLower.includes("oven"))
            productBonus += 200;
          // Penalize wrong categories
          if (pathLower.includes("welding") || pathLower.includes("electrode"))
            productBonus -= 200;
        }

        // Tumble Dryer Parts
        if (
          /\bdryer\b/.test(titleLower) &&
          (/\bvent\b/.test(titleLower) || /\bhose\b/.test(titleLower))
        ) {
          if (pathLower.includes("dryer") || pathLower.includes("washing"))
            productBonus += 150;
          if (leafLower.includes("part") || leafLower.includes("accessor"))
            productBonus += 50;
        }

        // Sliding Door Track (not barn door)
        if (
          /\bsliding\s*door\b/.test(titleLower) &&
          /\btrack\b/.test(titleLower)
        ) {
          if (pathLower.includes("door") && pathLower.includes("hardware"))
            productBonus += 200;
          // Penalize wrong categories
          if (pathLower.includes("photograph") || pathLower.includes("rail"))
            productBonus -= 200;
          if (pathLower.includes("vehicle") || pathLower.includes("car"))
            productBonus -= 150;
        }

        // Watches - Men's vs Women's
        if (/\bwatch\b/.test(titleLower) || /\bwristwatch\b/.test(titleLower)) {
          if (leafLower.includes("watch")) productBonus += 150;
          // Specific gender matching
          if (
            /\bwomen\b/.test(titleLower) ||
            /\bwoman\b/.test(titleLower) ||
            /\bwomen's\b/.test(titleLower)
          ) {
            if (pathLower.includes("women")) productBonus += 100;
            if (pathLower.includes("men") && !pathLower.includes("women"))
              productBonus -= 100;
          }
          if (
            /\bmen\b/.test(titleLower) ||
            /\bman\b/.test(titleLower) ||
            /\bmen's\b/.test(titleLower)
          ) {
            if (pathLower.includes("men") && !pathLower.includes("women"))
              productBonus += 100;
            if (pathLower.includes("women")) productBonus -= 100;
          }
        }

        // Filling Valve / Geberit / Toilet Parts
        if (
          /\bfilling\s*valve\b/.test(titleLower) ||
          /\bgeberit\b/.test(titleLower)
        ) {
          if (
            pathLower.includes("toilet") ||
            pathLower.includes("cistern") ||
            pathLower.includes("plumbing")
          ) {
            productBonus += 200;
          }
          if (leafLower.includes("valve") && pathLower.includes("toilet"))
            productBonus += 100;
          // Penalize wrong valve categories
          if (
            pathLower.includes("directional") ||
            pathLower.includes("hydraulic")
          )
            productBonus -= 200;
        }

        // Laptop Sleeve / Case / Bag
        if (
          /\blaptop\b/.test(titleLower) &&
          (/\bsleeve\b/.test(titleLower) ||
            /\bcase\b/.test(titleLower) ||
            /\bbag\b/.test(titleLower))
        ) {
          if (
            leafLower.includes("laptop") &&
            (leafLower.includes("bag") ||
              leafLower.includes("case") ||
              leafLower.includes("sleeve"))
          ) {
            productBonus += 250;
          }
          if (pathLower.includes("laptop")) productBonus += 80;
          // Penalize wrong categories
          if (pathLower.includes("cpap") || pathLower.includes("medical"))
            productBonus -= 200;
        }

        // Pasta Maker
        if (
          /\bpasta\s*maker\b/.test(titleLower) ||
          /\bpasta\s*machine\b/.test(titleLower)
        ) {
          if (leafLower.includes("pasta")) productBonus += 200;
          if (pathLower.includes("kitchen")) productBonus += 50;
        }

        // Rotary Tool
        if (/\brotary\s*tool\b/.test(titleLower)) {
          if (leafLower.includes("rotary")) productBonus += 200;
          if (pathLower.includes("power tool")) productBonus += 50;
        }

        // Beard Trimmer
        if (
          /\bbeard\s*trimmer\b/.test(titleLower) ||
          /\bstubble\b/.test(titleLower)
        ) {
          if (leafLower.includes("trimmer") || leafLower.includes("clipper"))
            productBonus += 150;
          if (pathLower.includes("shaving") || pathLower.includes("grooming"))
            productBonus += 50;
        }

        // Room Divider
        if (/\broom\s*divider\b/.test(titleLower)) {
          if (leafLower.includes("divider") || leafLower.includes("screen"))
            productBonus += 200;
          if (pathLower.includes("home decor")) productBonus += 50;
        }

        // Sauna
        if (/\bsauna\b/.test(titleLower)) {
          if (leafLower.includes("sauna")) productBonus += 200;
          // Penalize fitness clothing unless specifically a sauna suit
          if (
            pathLower.includes("fitness clothing") &&
            !/\bsuit\b/.test(titleLower)
          )
            productBonus -= 100;
        }

        // Photography Backdrop
        if (/\bbackdrop\b/.test(titleLower)) {
          if (
            leafLower.includes("backdrop") ||
            leafLower.includes("background")
          )
            productBonus += 200;
          if (pathLower.includes("photography") || pathLower.includes("studio"))
            productBonus += 50;
        }

        // Cassette Player
        if (/\bcassette\b/.test(titleLower)) {
          if (leafLower.includes("cassette") || leafLower.includes("portable"))
            productBonus += 150;
          if (pathLower.includes("portable audio")) productBonus += 50;
        }

        // Bed Alarm / Caregiver Pager
        if (
          /\bbed\s*alarm\b/.test(titleLower) ||
          /\bcaregiver\b/.test(titleLower)
        ) {
          if (pathLower.includes("medical") || pathLower.includes("monitoring"))
            productBonus += 150;
        }

        // Baby Sleeping Bag
        if (
          /\bbaby\s*sleeping\s*bag\b/.test(titleLower) ||
          /\bsleep\s*sack\b/.test(titleLower)
        ) {
          if (
            leafLower.includes("sleeping bag") ||
            leafLower.includes("sleepsack")
          )
            productBonus += 200;
          if (pathLower.includes("baby") || pathLower.includes("nursery"))
            productBonus += 80;
        }

        // Outdoor Lights
        if (/\boutdoor\b/.test(titleLower) && /\blight\b/.test(titleLower)) {
          if (pathLower.includes("outdoor") && pathLower.includes("lighting"))
            productBonus += 150;
          if (
            leafLower.includes("wall light") ||
            leafLower.includes("outdoor light")
          )
            productBonus += 100;
        }

        // Blackout Curtains
        if (/\bblackout\b/.test(titleLower) && /\bcurtain\b/.test(titleLower)) {
          if (leafLower.includes("curtain")) productBonus += 150;
          if (pathLower.includes("curtains")) productBonus += 50;
        }

        // Dishwasher Parts
        if (
          /\bdishwasher\b/.test(titleLower) &&
          (/\bbasket\b/.test(titleLower) || /\bcutlery\b/.test(titleLower))
        ) {
          if (pathLower.includes("dishwasher")) productBonus += 200;
          if (leafLower.includes("part") || leafLower.includes("accessor"))
            productBonus += 50;
        }

        // Dehumidifier
        if (/\bdehumidifier\b/.test(titleLower)) {
          if (leafLower.includes("dehumidifier")) productBonus += 200;
          if (pathLower.includes("air quality")) productBonus += 50;
        }

        // Water Bottle
        if (/\bwater\s*bottle\b/.test(titleLower)) {
          if (leafLower.includes("bottle") || leafLower.includes("water"))
            productBonus += 100;
        }

        // Mobility Aid / Stand Assist
        if (
          /\bmobility\s*aid\b/.test(titleLower) ||
          /\bstand\s*assist\b/.test(titleLower)
        ) {
          if (pathLower.includes("mobility") || pathLower.includes("medical"))
            productBonus += 150;
          if (leafLower.includes("mobility") || leafLower.includes("walking"))
            productBonus += 100;
        }

        // === LATEST PRODUCT BONUSES ===

        // Hand Dryer
        if (/\bhand\s*dryer\b/.test(titleLower)) {
          if (
            leafLower.includes("dryer") &&
            (pathLower.includes("bathroom") || pathLower.includes("commercial"))
          ) {
            productBonus += 200;
          }
          // Penalize wrong dryer categories
          if (
            pathLower.includes("hair") ||
            pathLower.includes("nail") ||
            pathLower.includes("tumble")
          ) {
            productBonus -= 150;
          }
        }

        // Bluetooth Headset / Wireless Headset
        if (
          /\bbluetooth\s*headset\b/.test(titleLower) ||
          /\bwireless\s*headset\b/.test(titleLower)
        ) {
          if (leafLower.includes("headphone") || leafLower.includes("headset"))
            productBonus += 200;
          if (
            pathLower.includes("portable audio") ||
            pathLower.includes("headphone")
          )
            productBonus += 80;
        }

        // WiFi Thermostat / Underfloor Heating
        if (/\bthermostat\b/.test(titleLower)) {
          if (leafLower.includes("thermostat")) productBonus += 150;
          if (pathLower.includes("heating") || pathLower.includes("hvac"))
            productBonus += 80;
          // Penalize industrial thermostats
          if (
            pathLower.includes("industrial") ||
            pathLower.includes("semiconductor")
          )
            productBonus -= 100;
        }
        if (
          /\bunderfloor\s*heating\b/.test(titleLower) ||
          /\bfloor\s*heating\b/.test(titleLower)
        ) {
          if (
            leafLower.includes("underfloor") ||
            leafLower.includes("heating system")
          )
            productBonus += 200;
          if (pathLower.includes("heating")) productBonus += 50;
        }

        // Hook-on High Chair / Portable High Chair
        if (
          /\bhook[\s-]?on\b/.test(titleLower) &&
          /\bchair\b/.test(titleLower)
        ) {
          if (
            leafLower.includes("high chair") ||
            leafLower.includes("highchair")
          )
            productBonus += 200;
          if (pathLower.includes("baby") || pathLower.includes("feeding"))
            productBonus += 80;
        }

        // Cat Repellent / Cat Scarer / Pest Control
        if (
          /\bcat\s*(repellent|scarer|deterrent)\b/.test(titleLower) ||
          /\banimal\s*repel/i.test(titleLower)
        ) {
          if (
            leafLower.includes("pest") ||
            leafLower.includes("repel") ||
            leafLower.includes("ultrasonic")
          ) {
            productBonus += 200;
          }
          if (
            pathLower.includes("pest control") ||
            pathLower.includes("garden")
          )
            productBonus += 80;
          // Penalize wrong categories
          if (pathLower.includes("book") || pathLower.includes("magazine"))
            productBonus -= 200;
        }

        // TV Remote / Smart TV Remote
        if (
          /\btv\s*remote\b/.test(titleLower) ||
          /\bsmart\s*tv\s*remote\b/.test(titleLower) ||
          (/\blg\b/.test(titleLower) && /\bremote\b/.test(titleLower))
        ) {
          if (
            leafLower.includes("remote control") ||
            leafLower.includes("remote")
          )
            productBonus += 200;
          if (pathLower.includes("tv") || pathLower.includes("sound vision"))
            productBonus += 80;
          // Penalize wrong remote categories
          if (
            pathLower.includes("vehicle") ||
            pathLower.includes("hoist") ||
            pathLower.includes("lifting")
          ) {
            productBonus -= 150;
          }
        }

        // Passport Cover / Passport Holder
        if (/\bpassport\s*(cover|holder|case)\b/.test(titleLower)) {
          if (leafLower.includes("passport") || leafLower.includes("id holder"))
            productBonus += 250;
          if (pathLower.includes("travel") || pathLower.includes("luggage"))
            productBonus += 80;
        }

        // === MORE LATEST PRODUCT BONUSES ===

        // Barn Door Hardware (stronger)
        if (
          /\bbarn\s*door\b/.test(titleLower) ||
          /\bsliding\s*barn\b/.test(titleLower)
        ) {
          if (
            pathLower.includes("door") &&
            (pathLower.includes("hardware") || pathLower.includes("handle"))
          ) {
            productBonus += 200;
          }
          // Penalize wrong categories
          if (
            pathLower.includes("vehicle") ||
            pathLower.includes("photograph") ||
            pathLower.includes("antique")
          ) {
            productBonus -= 150;
          }
        }

        // Baby Play Mat / Floor Mat / Foam Mat
        if (
          /\bplay\s*mat\b/.test(titleLower) ||
          /\bplaymat\b/.test(titleLower) ||
          /\bfoam\s*mat\b/.test(titleLower)
        ) {
          if (leafLower.includes("playmat") || leafLower.includes("play mat"))
            productBonus += 250;
          if (pathLower.includes("baby") || pathLower.includes("toys"))
            productBonus += 80;
          // If baby mentioned, prioritize baby category
          if (/\bbaby\b/.test(titleLower)) {
            if (pathLower.includes("baby")) productBonus += 100;
          }
        }
        if (/\bfloor\s*mat\b/.test(titleLower) && /\bbaby\b/.test(titleLower)) {
          if (pathLower.includes("baby")) productBonus += 150;
        }

        // Caster Wheels
        if (/\bcaster\b/.test(titleLower) || /\bcasters\b/.test(titleLower)) {
          if (leafLower.includes("caster") || leafLower.includes("wheel"))
            productBonus += 200;
          if (pathLower.includes("material handling")) productBonus += 80;
        }

        // Decorative Box / Glass Box / Trinket Box
        if (
          /\bdecorative\s*box\b/.test(titleLower) ||
          /\bglass.*box\b/.test(titleLower) ||
          /\btrinket\s*box\b/.test(titleLower)
        ) {
          if (leafLower.includes("box") || leafLower.includes("trinket"))
            productBonus += 150;
          if (
            pathLower.includes("storage") ||
            pathLower.includes("decorative") ||
            pathLower.includes("collectables")
          )
            productBonus += 50;
        }
        if (
          /\blidded\s*box\b/.test(titleLower) ||
          /\bvintage.*box\b/.test(titleLower)
        ) {
          if (leafLower.includes("box") || leafLower.includes("storage"))
            productBonus += 100;
        }

        // Drain Rods / Drain Cleaning
        if (
          /\bdrain\s*rod\b/.test(titleLower) ||
          /\bdrain\s*rods\b/.test(titleLower) ||
          /\bdrain\s*cleaning\b/.test(titleLower)
        ) {
          if (
            leafLower.includes("drain") ||
            pathLower.includes("pipe tool") ||
            pathLower.includes("plumbing")
          ) {
            productBonus += 200;
          }
        }

        // Footmuff / Cosytoes
        if (
          /\bfootmuff\b/.test(titleLower) ||
          /\bcosytoes?\b/.test(titleLower)
        ) {
          if (
            leafLower.includes("footmuff") ||
            leafLower.includes("cosytoes") ||
            leafLower.includes("apron")
          ) {
            productBonus += 250;
          }
          if (pathLower.includes("pushchair") || pathLower.includes("pram"))
            productBonus += 80;
        }

        // Galaxy Z Fold Case / Samsung Case
        if (
          /\bgalaxy\s*z\s*fold\b/.test(titleLower) ||
          /\bz\s*fold\b/.test(titleLower)
        ) {
          if (
            leafLower.includes("cases") ||
            leafLower.includes("covers") ||
            leafLower.includes("skins")
          ) {
            productBonus += 200;
          }
          if (pathLower.includes("mobile") || pathLower.includes("phone"))
            productBonus += 80;
        }

        // === STRONG CATEGORY PROTECTION PENALTIES ===

        // Prevent AUTOMOTIVE products going to TOYS
        if (
          /\btyre\b/.test(titleLower) ||
          /\btire\b/.test(titleLower) ||
          /\brim\b/.test(titleLower) ||
          /\brims\b/.test(titleLower) ||
          /\bwheel\b/.test(titleLower)
        ) {
          if (
            pathLower.includes("toy") ||
            pathLower.includes("game") ||
            pathLower.includes("hobbies")
          ) {
            productBonus -= 500; // MASSIVE penalty
          }
          if (
            pathLower.includes("vehicle") ||
            pathLower.includes("car") ||
            pathLower.includes("automotive") ||
            pathLower.includes("motor")
          ) {
            productBonus += 300; // Strong bonus for correct category
          }
        }

        // Prevent VEHICLE PARTS going to TOYS
        if (
          /\bcar\s*part\b/.test(titleLower) ||
          /\bvehicle\b/.test(titleLower) ||
          /\bautomotive\b/.test(titleLower)
        ) {
          if (
            pathLower.includes("toy") ||
            pathLower.includes("game") ||
            pathLower.includes("hobbies") ||
            pathLower.includes("collectables")
          ) {
            productBonus -= 500;
          }
          if (
            pathLower.includes("vehicle") ||
            pathLower.includes("automotive") ||
            pathLower.includes("car parts")
          ) {
            productBonus += 300;
          }
        }

        // Prevent TOOLS going to TOYS
        if (
          /\bdrill\b/.test(titleLower) ||
          /\bsaw\b/.test(titleLower) ||
          /\bhammer\b/.test(titleLower) ||
          /\bwrench\b/.test(titleLower) ||
          /\bscrewdriver\b/.test(titleLower)
        ) {
          if (pathLower.includes("toy") || pathLower.includes("game")) {
            productBonus -= 400;
          }
          if (
            pathLower.includes("tool") ||
            pathLower.includes("diy") ||
            pathLower.includes("hardware")
          ) {
            productBonus += 200;
          }
        }

        // Prevent ELECTRONICS going to TOYS (unless it's actually a toy)
        if (!/\btoy\b/.test(titleLower)) {
          if (
            /\bphone\b/.test(titleLower) ||
            /\btablet\b/.test(titleLower) ||
            /\blaptop\b/.test(titleLower) ||
            /\bcamera\b/.test(titleLower)
          ) {
            if (
              pathLower.includes("toy") ||
              (pathLower.includes("game") && !pathLower.includes("video game"))
            ) {
              productBonus -= 400;
            }
          }
        }

        // Prevent BABY items going to wrong categories
        if (
          /\bbaby\b/.test(titleLower) ||
          /\binfant\b/.test(titleLower) ||
          /\bnewborn\b/.test(titleLower)
        ) {
          if (pathLower.includes("baby") || pathLower.includes("nursery")) {
            productBonus += 150;
          }
        }

        // Prevent PET products going to wrong categories
        if (
          /\bdog\b/.test(titleLower) ||
          /\bcat\b/.test(titleLower) ||
          /\bpet\b/.test(titleLower)
        ) {
          if (
            pathLower.includes("pet") ||
            pathLower.includes("dog") ||
            pathLower.includes("cat")
          ) {
            productBonus += 150;
          }
          if (pathLower.includes("toy") && !pathLower.includes("pet")) {
            productBonus -= 200;
          }
        }

        // Penalize wrong matches
        if (
          leafLower.includes("air tool") ||
          leafLower.includes("air suspension") ||
          leafLower.includes("air intake")
        ) {
          productBonus -= 100;
        }

        // Coverage bonus
        const origMatched = [...scoreData.matched].filter((m) =>
          titleTokens.includes(m.replace("*", ""))
        ).length;
        const coverageBonus =
          (origMatched / Math.max(titleTokens.length, 1)) * 30;

        const totalScore =
          leafScore +
          pathScore +
          depthBonus +
          domainBonus +
          productBonus +
          coverageBonus +
          excelCategoryBonus +
          googleBonus;

        results.push({
          categoryId: cat.id,
          categoryPath: cat.path,
          score: totalScore,
          leafScore: leafScore,
          domainBonus: domainBonus,
          productBonus: productBonus,
          excelCategoryBonus: excelCategoryBonus,
          matched: scoreData.matched,
          depth: cat.depth,
        });
      });

      // Sort ALL results by total score
      results.sort((a, b) => b.score - a.score);

      if (results.length === 0 || results[0].score < 5) {
        return {
          categoryId: "47155",
          categoryPath: "Default",
          score: 0,
          confidence: "none",
        };
      }

      const best = results[0];
      let confidence = "low";
      if (best.score >= 300) confidence = "high";
      else if (best.score >= 150) confidence = "medium";

      console.log(
        `Category: "${best.categoryPath}" | Score: ${best.score.toFixed(
          0
        )} | Confidence: ${confidence}`
      );

      return {
        categoryId: best.categoryId,
        categoryPath: best.categoryPath,
        score: best.score,
        confidence: confidence,
      };
    }
  }

  const categoryMatcher = new IntelligentCategoryMatcher();

  // =============================================================================
  // UTILITY FUNCTIONS
  // =============================================================================

  function detectColumnMapping(excelHeaders) {
    const mapping = {
      brand: -1,
      colour: -1,
      material: -1,
      size: -1,
      weight: -1,
      condition: -1,
    };

    console.log("=== DETECTING COLUMN MAPPING ===");
    console.log("Excel Headers:", excelHeaders);

    excelHeaders.forEach((header, idx) => {
      const headerLower = String(header || "")
        .toLowerCase()
        .trim();

      // Brand detection - check for exact "brand" or contains "brand"
      if (headerLower === "brand" || headerLower.includes("brand")) {
        mapping.brand = idx;
        console.log(`✓ BRAND column found at index ${idx}: "${header}"`);
      }
      if (headerLower.includes("colour") || headerLower.includes("color")) {
        mapping.colour = idx;
        console.log(`✓ COLOUR column found at index ${idx}: "${header}"`);
      }
      if (headerLower.includes("material")) {
        mapping.material = idx;
        console.log(`✓ MATERIAL column found at index ${idx}: "${header}"`);
      }
      if (headerLower.includes("size") && !headerLower.includes("pack")) {
        mapping.size = idx;
        console.log(`✓ SIZE column found at index ${idx}: "${header}"`);
      }
      if (headerLower.includes("weight")) {
        mapping.weight = idx;
        console.log(`✓ WEIGHT column found at index ${idx}: "${header}"`);
      }
      if (headerLower.includes("condition")) {
        mapping.condition = idx;
        console.log(`✓ CONDITION column found at index ${idx}: "${header}"`);
      }
    });

    console.log("Final Column Mapping:", mapping);

    if (mapping.brand === -1) {
      console.warn(
        "⚠️ WARNING: No Brand column detected! Check your Excel header names."
      );
    }

    return mapping;
  }

  const round = (num, decimals = 2) => {
    return Math.round(num * Math.pow(10, decimals)) / Math.pow(10, decimals);
  };

  const getExcelColumn = (col) => {
    let temp,
      letter = "";
    while (col >= 0) {
      temp = col % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      col = (col - temp - 1) / 26;
    }
    return letter;
  };

  const formatDateToSKU = (dateStr) => {
    let year, month, day;

    if (dateStr.includes("-")) {
      [year, month, day] = dateStr.split("-");
    } else if (dateStr.length === 8) {
      year = dateStr.substring(0, 4);
      month = dateStr.substring(4, 6);
      day = dateStr.substring(6, 8);
    } else {
      return "010101";
    }

    return `${day}${month}${year.slice(2)}`;
  };

  const shortenTitle = (title, maxLength = 70) => {
    if (!title) return "";
    const fillers = ["with", "for", "the", "and", "&", "-", "a", "an"];
    const words = title.split(/[\s,]+/);
    let result = [];
    let length = 0;

    for (let word of words) {
      const cleanWord = word.replace(/[^\w\s]/g, "");
      if (fillers.includes(cleanWord.toLowerCase())) continue;
      if (length + word.length + 1 <= maxLength) {
        result.push(word);
        length += word.length + 1;
      } else {
        break;
      }
    }
    return result.join(" ").substring(0, maxLength).trim();
  };

  function showStatus(message, type) {
    const status = document.getElementById("status");
    status.textContent = message;
    status.className = `status ${type}`;
    status.style.display = "block";

    if (type === "success") {
      setTimeout(() => (status.style.display = "none"), 5000);
    }
  }

  function isDescriptionEmpty(desc) {
    if (!desc) return true;
    const cleaned = String(desc).trim().toUpperCase();
    return (
      cleaned === "N/A" ||
      cleaned === "" ||
      cleaned === "NA" ||
      cleaned === "-" ||
      cleaned === "NONE"
    );
  }

  // =============================================================================
  // PDF PROCESSING
  // =============================================================================

  async function processPDFs(pdfFiles) {
    const pdfDataArray = [];

    for (let i = 0; i < pdfFiles.length; i++) {
      const file = pdfFiles[i];
      showStatus(`Reading PDF ${i + 1}/${pdfFiles.length}...`, "info");

      try {
        const arrayBuffer = await file.arrayBuffer();
        const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer });
        const pdfDoc = await loadingTask.promise;

        let pdfText = "";
        for (let pageNum = 1; pageNum <= pdfDoc.numPages; pageNum++) {
          const page = await pdfDoc.getPage(pageNum);
          const textContent = await page.getTextContent();
          const pageText = textContent.items.map((item) => item.str).join(" ");
          pdfText += pageText + "\n";
        }

        const shippingMatch =
          pdfText.match(/Net\s+shipping.*?£\s*(\d+\.?\d*)/i) ||
          pdfText.match(/shipping.*?£\s*(\d+\.?\d*)/i);
        const shipping = shippingMatch ? parseFloat(shippingMatch[1]) : 0;

        const invoiceDateMatch = pdfText.match(
          /ORDER\s+DATE[:\s]+(\d{2}[-/]\d{2}[-/]\d{4})/i
        );
        const invoiceDate = invoiceDateMatch ? invoiceDateMatch[1] : "";

        const vendorNumberMatch = pdfText.match(/INVOICE\s+([A-Z0-9]+)/i);
        const vendorNumber = vendorNumberMatch ? vendorNumberMatch[1] : "";

        pdfDataArray.push({
          text: pdfText,
          shipping: shipping,
          name: file.name,
          index: i,
          invoiceDate: invoiceDate,
          vendorNumber: vendorNumber,
        });
      } catch (error) {
        console.error(`Error processing PDF ${file.name}:`, error);
        showStatus(`Error reading PDF: ${file.name}`, "error");
      }
    }

    return pdfDataArray;
  }

  // =============================================================================
  // FILE HANDLING
  // =============================================================================

  async function handleExcelFile(file) {
    showStatus("Reading Excel file...", "info");

    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: "array" });
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      excelData = XLSX.utils.sheet_to_json(firstSheet, {
        header: 1,
        defval: "",
      });

      showStatus(`Excel file loaded: ${excelData.length - 1} rows`, "success");
      document.getElementById("fileInfo").textContent = `📊 ${file.name} (${
        excelData.length - 1
      } rows)`;
    } catch (error) {
      showStatus("Error reading Excel file", "error");
      console.error(error);
    }
  }

  async function handlePDFFiles(files) {
    showStatus("Processing PDFs...", "info");
    pdfDataArray = await processPDFs(files);

    const totalShipping = pdfDataArray.reduce(
      (sum, pdf) => sum + pdf.shipping,
      0
    );
    document.getElementById("pdfInfo").textContent = `📄 ${
      files.length
    } PDF(s) | £${totalShipping.toFixed(2)} total shipping`;
    showStatus("PDFs processed successfully", "success");
  }

  async function handleCategoryFile(file) {
    showStatus("Reading Category Map...", "info");

    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: "array" });
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      categoryMap = XLSX.utils.sheet_to_json(firstSheet);

      categoryMatcher.build(categoryMap);
      const validCount = categoryMatcher.categories.length;

      if (validCount === 0) {
        showStatus(`⚠️ No valid categories found`, "error");
      } else {
        showStatus(`Category Map: ${validCount} categories loaded`, "success");
      }

      document.getElementById(
        "categoryInfo"
      ).textContent = `🗂️ ${file.name} (${validCount} categories)`;
    } catch (error) {
      showStatus("Error reading Category Map", "error");
      console.error(error);
    }
  }

  // =============================================================================
  // EVENT LISTENERS
  // =============================================================================

  document.getElementById("excelFile").addEventListener("change", async (e) => {
    if (e.target.files.length > 0) {
      await handleExcelFile(e.target.files[0]);
    }
  });

  document.getElementById("pdfFiles").addEventListener("change", async (e) => {
    if (e.target.files.length > 0 && excelData.length > 0) {
      await handlePDFFiles(Array.from(e.target.files));
    } else if (excelData.length === 0) {
      showStatus("Please upload Excel file first", "error");
    }
  });

  document
    .getElementById("categoryFile")
    .addEventListener("change", async (e) => {
      if (e.target.files.length > 0) {
        await handleCategoryFile(e.target.files[0]);
      }
    });

  document.getElementById("processBtn").addEventListener("click", async () => {
    if (excelData.length === 0) {
      showStatus("Please upload an Excel file first", "error");
      return;
    }

    const dateInput = document.getElementById("dateInput").value;
    if (!dateInput) {
      showStatus("Please enter a date (YYYY-MM-DD or YYYYMMDD)", "error");
      return;
    }

    const isValidFormat =
      /^\d{4}-\d{2}-\d{2}$/.test(dateInput) || /^\d{8}$/.test(dateInput);
    if (!isValidFormat) {
      showStatus("Invalid date format. Use YYYY-MM-DD or YYYYMMDD", "error");
      return;
    }

    showStatus("Processing files...", "info");

    try {
      await processFiles(dateInput);
    } catch (error) {
      showStatus("Error processing files: " + error.message, "error");
      console.error(error);
    }
  });

  // =============================================================================
  // MAIN PROCESSING FUNCTION - FIXED SKU FORMAT WITH COST PER ONE
  // =============================================================================

  async function processFiles(date) {
    const dataRows = excelData.slice(1);

    // Extract SKUs from Column C (index 2)
    showStatus("Extracting SKUs from Excel...", "info");
    const manifestSKUs = [];

    for (let idx = 0; idx < dataRows.length; idx++) {
      const row = dataRows[idx];
      const sku = (row[2] && String(row[2]).trim()) || "";
      if (sku) {
        manifestSKUs.push({ sku, rowIndex: idx });
      }
    }

    // === PRODUCT LOOKUP DISABLED ===
    // API-based product lookup has been removed to avoid rate limits.
    // Category matching will use title-based matching only.
    const productLookupData = new Map(); // Empty map for compatibility

    // Search PDFs for prices
    showStatus("Searching PDFs for SKU prices...", "info");
    const items = [];

    for (let i = 0; i < manifestSKUs.length; i++) {
      const { sku } = manifestSKUs[i];
      const skuRegex = new RegExp(`\\b${sku.replace(/[-]/g, "[-]?")}\\b`, "gi");
      let found = false;

      for (let pdfData of pdfDataArray) {
        const match = skuRegex.exec(pdfData.text);

        if (match) {
          const skuIndex = match.index;
          const textAfterSKU = pdfData.text.substring(skuIndex, skuIndex + 200);
          const priceMatches = textAfterSKU.match(/£\s*(\d+\.?\d*)/g);

          if (priceMatches && priceMatches.length > 0) {
            const firstPrice = parseFloat(priceMatches[0].replace(/£\s*/g, ""));
            items.push({
              sku: sku,
              cost: firstPrice,
              pdfIndex: pdfData.index,
              shipping: pdfData.shipping,
              invoiceDate: pdfData.invoiceDate,
              vendorNumber: pdfData.vendorNumber,
            });
            found = true;
            break;
          }
        }
      }

      if (!found) {
        items.push({
          sku: sku,
          cost: 0,
          pdfIndex: -1,
          shipping: 0,
          invoiceDate: "",
          vendorNumber: "",
        });
      }
    }

    showStatus("Calculating costs and shipping...", "info");

    // Calculate proportional costs for duplicate SKUs
    const skuGroups = {};
    manifestSKUs.forEach(({ sku, rowIndex }) => {
      if (!skuGroups[sku]) skuGroups[sku] = [];
      skuGroups[sku].push(rowIndex);
    });

    const proportionalCosts = {};
    for (const sku of Object.keys(skuGroups)) {
      const rowIndices = skuGroups[sku];
      const pdfItem = items.find((item) => item.sku === sku);
      const invoiceCost = pdfItem ? pdfItem.cost : 0;

      if (rowIndices.length > 1) {
        const totalRRP = rowIndices.reduce((sum, idx) => {
          const rrp = parseFloat(dataRows[idx][22]) || 0;
          return sum + rrp;
        }, 0);

        rowIndices.forEach((idx) => {
          const itemRRP = parseFloat(dataRows[idx][22]) || 0;
          const proportionalCost =
            totalRRP > 0 ? (itemRRP / totalRRP) * invoiceCost : 0;
          proportionalCosts[idx] = proportionalCost;
        });
      } else {
        proportionalCosts[rowIndices[0]] = invoiceCost;
      }
    }

    // Allocate shipping by weight × quantity
    const itemsByPDF = {};

    dataRows.forEach((row, idx) => {
      const manifestSKU = manifestSKUs[idx]?.sku || "";
      const pdfItem = items.find((item) => item.sku === manifestSKU);
      const pdfIndex = pdfItem ? pdfItem.pdfIndex : -1;
      const pdfShipping = pdfItem ? pdfItem.shipping : 0;

      if (!itemsByPDF[pdfIndex]) {
        itemsByPDF[pdfIndex] = {
          items: [],
          shipping: pdfShipping,
          totalCurrency: 0,
        };
      }

      const weight = parseFloat(row[20]) || 0;
      const quantity = parseInt(row[17]) || 1;
      const currency = weight * quantity;

      itemsByPDF[pdfIndex].items.push({ rowIndex: idx, currency });
      itemsByPDF[pdfIndex].totalCurrency += currency;
    });

    const itemShipping = {};
    for (const pdfIndex of Object.keys(itemsByPDF)) {
      const pdfGroup = itemsByPDF[pdfIndex];

      pdfGroup.items.forEach(({ rowIndex, currency }) => {
        const shipping =
          pdfGroup.totalCurrency > 0
            ? (currency / pdfGroup.totalCurrency) * pdfGroup.shipping
            : 0;
        itemShipping[rowIndex] = shipping;
      });
    }

    showStatus("Extracting item specifics and matching categories...", "info");

    const dateSKU = formatDateToSKU(date);
    const asinToSKU = {};
    let skuCounter = 1;

    const headerRow = [
      "Order Date",
      "Vendor",
      ...excelData[0].slice(0, 24),
      "Cost",
      "Shipping ",
      "VAT",
      "Total Cost",
      "Cost Per One",
      "SKU",
      "Sales Price",
      "Category ID",
      "Title",
    ];

    const processedRows = [headerRow];
    const columnMapping = detectColumnMapping(excelData[0]);

    // Log the Brand column detection result
    console.log("=== BRAND COLUMN DETECTION ===");
    console.log("Input Excel Headers (first 15):", excelData[0].slice(0, 15));
    console.log("Detected Brand column index:", columnMapping.brand);
    if (columnMapping.brand >= 0) {
      console.log("Brand header name:", excelData[0][columnMapping.brand]);
      console.log("Sample Brand values from first 3 data rows:");
      for (let i = 0; i < Math.min(3, dataRows.length); i++) {
        console.log(`  Row ${i}: "${dataRows[i][columnMapping.brand]}"`);
      }
    }

    const allSpecificsKeys = new Set();
    const itemSpecificsArray = [];
    const allDescriptions = [];

    // FIRST PASS: Calculate all costs to get Cost Per One for SKU
    const costPerOneArray = [];

    for (let idx = 0; idx < dataRows.length; idx++) {
      const row = dataRows[idx];
      const quantity = parseInt(row[17]) || 1;

      const cost = round(proportionalCosts[idx] || 0);
      const shipping = round(itemShipping[idx] || 0);
      const vat = round((cost + shipping) * 0.2);
      const totalCostCalc = round(cost + shipping + vat);
      const costPerOne = round(
        quantity > 0 ? totalCostCalc / quantity : totalCostCalc
      );

      costPerOneArray.push(costPerOne);
    }

    // Extract specifics and generate descriptions for ALL items
    for (let idx = 0; idx < dataRows.length; idx++) {
      const row = dataRows[idx];
      const title = row[3] || "";
      const originalDescription = row[6] || "";

      // Extract specifics (includes Type detection)
      const specifics = enhancedSpecificsGenerator.extractAllSpecifics(
        title,
        originalDescription
      );

      // PRIORITY: Use Brand from Excel column if available - THIS OVERRIDES ANY PATTERN-DETECTED BRAND
      // In INPUT Excel: Column I = index 8 is the Brand column
      const BRAND_COLUMN_INDEX = 8; // Hardcoded fallback: Column I (index 8)

      let brandColumnIdx =
        columnMapping.brand >= 0 ? columnMapping.brand : BRAND_COLUMN_INDEX;

      if (brandColumnIdx >= 0) {
        const excelBrand = String(row[brandColumnIdx] || "").trim();
        if (
          excelBrand &&
          excelBrand !== "" &&
          excelBrand.toLowerCase() !== "undefined" &&
          excelBrand.toLowerCase() !== "null"
        ) {
          specifics.Brand = excelBrand;
          if (idx < 5) {
            // Log first 5 items for debugging
            console.log(
              `Row ${idx}: Brand from Excel column [${brandColumnIdx}] = "${excelBrand}"`
            );
          }
        } else {
          if (idx < 5) {
            console.log(
              `Row ${idx}: Brand column [${brandColumnIdx}] is empty. Row data:`,
              row.slice(0, 12)
            );
            console.log(
              `Row ${idx}: Using pattern-detected brand: "${
                specifics.Brand || "NONE"
              }"`
            );
          }
        }
      } else {
        if (idx < 5) {
          console.log(
            `Row ${idx}: No Brand column mapped, using pattern-detected: "${
              specifics.Brand || "NONE"
            }"`
          );
        }
      }

      // Also check for other column-based specifics
      if (columnMapping.colour >= 0 && row[columnMapping.colour]) {
        const excelColour = String(row[columnMapping.colour]).trim();
        if (excelColour && excelColour !== "") {
          specifics.Color = excelColour;
        }
      }
      if (columnMapping.material >= 0 && row[columnMapping.material]) {
        const excelMaterial = String(row[columnMapping.material]).trim();
        if (excelMaterial && excelMaterial !== "") {
          specifics.Material = excelMaterial;
        }
      }
      if (columnMapping.size >= 0 && row[columnMapping.size]) {
        const excelSize = String(row[columnMapping.size]).trim();
        if (excelSize && excelSize !== "") {
          specifics.Size = excelSize;
        }
      }

      // USE SUBCATEGORY (Column K, index 10) AS TYPE - PRIORITY OVER PATTERN DETECTION
      const SUBCATEGORY_COLUMN_INDEX = 10;
      const excelSubCategory = String(
        row[SUBCATEGORY_COLUMN_INDEX] || ""
      ).trim();
      if (
        excelSubCategory &&
        excelSubCategory !== "" &&
        excelSubCategory.toLowerCase() !== "undefined"
      ) {
        specifics.Type = excelSubCategory;
        if (idx < 5) {
          console.log(
            `Row ${idx}: Type from SubCategory column [${SUBCATEGORY_COLUMN_INDEX}] = "${excelSubCategory}"`
          );
        }
      }

      // STORE CATEGORY (Column J, index 9) for category matching (internal use only, NOT in output)
      const CATEGORY_COLUMN_INDEX = 9;
      const excelCategory = String(row[CATEGORY_COLUMN_INDEX] || "").trim();
      // Store in temporary variables for matching, but NOT in specifics object
      const _matchingExcelCategory = excelCategory;
      const _matchingExcelSubCategory = excelSubCategory;

      // === PRODUCT LOOKUP DISABLED ===
      // API-based specifics extraction removed. Using Excel data only.

      itemSpecificsArray.push(specifics);

      // Add all specific keys (Type is now included) - but exclude internal keys
      Object.keys(specifics).forEach((key) => {
        // Don't add keys starting with underscore to the output
        if (!key.startsWith("_")) {
          allSpecificsKeys.add(key);
        }
      });

      // Generate or use existing description
      let finalDescription = originalDescription;
      if (isDescriptionEmpty(originalDescription)) {
        finalDescription = descriptionGenerator.generateDescription(
          title,
          specifics.Brand || "",
          specifics.Type || "General Product",
          specifics
        );
      }
      allDescriptions.push(finalDescription);
    }

    // Ensure core specifics are always included (even if empty for some items)
    const coreSpecifics = ["Brand", "Type", "Color", "Material", "Size"];
    coreSpecifics.forEach((key) => allSpecificsKeys.add(key));

    const sortedSpecificsKeys = Array.from(allSpecificsKeys).sort();

    // Build processed rows with NEW SKU FORMAT: DateSKU/RoundedCostPerOne
    for (let idx = 0; idx < dataRows.length; idx++) {
      const row = dataRows[idx];

      const asin = String(row[5] || "").toLowerCase();
      const weight = parseFloat(row[20]) || 0;
      const quantity = parseInt(row[17]) || 1;
      const unitRRP = parseFloat(row[22]) || 0;
      const originalTitle = row[3] || "";

      const manifestSKU = manifestSKUs[idx]?.sku || "";
      const pdfItem = items.find((item) => item.sku === manifestSKU);

      const invoiceDate = pdfItem ? pdfItem.invoiceDate : "";
      const vendorNumber = pdfItem ? pdfItem.vendorNumber : "";

      const cost = round(proportionalCosts[idx] || 0);
      const shipping = round(itemShipping[idx] || 0);

      // Generate base SKU from ASIN
      if (asin && !asinToSKU[asin]) {
        asinToSKU[asin] = dateSKU + String(skuCounter).padStart(2, "0");
        skuCounter++;
      }
      const baseSKU = asinToSKU[asin] || dateSKU + "01";

      // NEW: Add rounded COST PER ONE to SKU (not unit RRP)
      const costPerOne = costPerOneArray[idx];
      const roundedCostPerOne = Math.ceil(costPerOne);
      const fullSKU = `${baseSKU}/${roundedCostPerOne}`;

      const vat = round((cost + shipping) * 0.2);
      const totalCostCalc = round(cost + shipping + vat);

      const salesPrice = round(unitRRP * 0.82);

      const shortenedTitle = shortenTitle(originalTitle);
      const fullTitle = `${shortenedTitle} RRP £${Math.ceil(unitRRP)}`;

      // Get Category and SubCategory from Excel to improve matching
      const CATEGORY_COLUMN_INDEX = 9;
      const SUBCATEGORY_COLUMN_INDEX = 10;
      const excelCategory = String(row[CATEGORY_COLUMN_INDEX] || "").trim();
      const excelSubCategory = String(
        row[SUBCATEGORY_COLUMN_INDEX] || ""
      ).trim();

      // Use INTELLIGENT category matcher - pass Excel category info and rowIndex for Google lookup
      const categoryResult = categoryMatcher.match(
        originalTitle,
        row[4] || "",
        excelCategory,
        excelSubCategory,
        idx
      );
      const categoryId = categoryResult.categoryId;

      const rowData = [];
      for (let i = 0; i < 24; i++) {
        if (i === 21) {
          rowData.push(weight * quantity);
        } else {
          rowData.push(row[i]);
        }
      }

      const newRow = [
        invoiceDate,
        vendorNumber,
        ...rowData,
        cost,
        shipping,
        vat,
        totalCostCalc,
        costPerOne,
        fullSKU, // Now includes /CostPerOne
        salesPrice,
        categoryId,
        fullTitle,
      ];

      processedRows.push(newRow);
    }

    // Build Excel file with formulas
    showStatus("Building Excel file with formulas...", "info");

    const newWorkbook = XLSX.utils.book_new();
    const newWorksheet = XLSX.utils.aoa_to_sheet(processedRows);

    for (let i = 0; i < dataRows.length; i++) {
      const excelRow = i + 2;

      const colQuantity = getExcelColumn(19);
      const colUnitRRP = getExcelColumn(24);
      const colCost = getExcelColumn(26);
      const colShipping = getExcelColumn(27);
      const colVAT = getExcelColumn(28);
      const colTotalCost = getExcelColumn(29);
      const colCostPerOne = getExcelColumn(30);
      const colSalesPrice = getExcelColumn(32);

      // VAT Formula
      const vatFormula = `=ROUND((${colCost}${excelRow}+${colShipping}${excelRow})*0.2,2)`;
      newWorksheet[`${colVAT}${excelRow}`] = {
        t: "n",
        f: vatFormula,
        v: processedRows[i + 1][28],
      };

      // Total Cost Formula
      const totalCostFormula = `=ROUND(${colCost}${excelRow}+${colShipping}${excelRow}+${colVAT}${excelRow},2)`;
      newWorksheet[`${colTotalCost}${excelRow}`] = {
        t: "n",
        f: totalCostFormula,
        v: processedRows[i + 1][29],
      };

      // Cost Per One Formula
      const costPerOneFormula = `=ROUND(${colTotalCost}${excelRow}/${colQuantity}${excelRow},2)`;
      newWorksheet[`${colCostPerOne}${excelRow}`] = {
        t: "n",
        f: costPerOneFormula,
        v: processedRows[i + 1][30],
      };

      // Sales Price Formula
      const salesPriceFormula = `=ROUND(${colUnitRRP}${excelRow}*0.82,2)`;
      newWorksheet[`${colSalesPrice}${excelRow}`] = {
        t: "n",
        f: salesPriceFormula,
        v: processedRows[i + 1][32],
      };
    }

    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, dateSKU);

    const wbout = XLSX.write(newWorkbook, { bookType: "xlsx", type: "array" });
    const excelBlob = new Blob([wbout], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const excelUrl = URL.createObjectURL(excelBlob);

    document.getElementById("downloadInternal").href = excelUrl;
    document.getElementById(
      "downloadInternal"
    ).download = `LISTING_${dateSKU}.xlsx`;
    document.getElementById("downloadInternal").style.display = "inline-block";

    showStatus("Excel file ready. Creating eBay CSV...", "info");

    // Store data for eBay CSV creation
    const ebayData = {
      dateSKU,
      processedRows,
      itemSpecificsArray,
      sortedSpecificsKeys,
      allDescriptions,
      columnMapping,
    };

    // Create eBay CSV
    createEbayCSV(ebayData);
  }

  // =============================================================================
  // CREATE EBAY CSV WITH C: PREFIX FOR ALL SPECIFICS
  // =============================================================================

  function createEbayCSV(ebayData) {
    showStatus("Creating eBay CSV with item specifics...", "info");

    const ebayRows = [];

    // Add #INFO headers
    ebayRows.push(
      "#INFO,Version=0.0.2,Template= eBay-draft-listings-template_GB,,,,,,,,"
    );
    ebayRows.push(
      "#INFO Action and Category ID are required fields. 1) Set Action to Draft 2) Please find the category ID for your listings here: https://pages.ebay.com/sellerinformation/news/categorychanges.html,,,,,,,,,,"
    );
    ebayRows.push(
      "#INFO After you've successfully uploaded your draft from the Seller Hub Reports tab, complete your drafts to active listings here: https://www.ebay.co.uk/sh/lst/drafts,,,,,,,,,,"
    );
    ebayRows.push("#INFO,,,,,,,,,,");

    // Build header row with C: prefix for ALL item specifics
    const headerBase = [
      "Action(SiteID=UK|Country=GB|Currency=GBP|Version=1193|CC=UTF-8)",
      "Custom label (SKU)",
      "Category ID",
      "Title",
      "UPC",
      "Price",
      "Quantity",
      "Item photo URL",
      "Condition ID",
      "Description",
      "Format",
    ];

    // Add C: prefix to ALL specifics columns
    const headerSpecifics = ebayData.sortedSpecificsKeys.map(
      (key) => `C:${key}`
    );
    const fullHeader = [...headerBase, ...headerSpecifics];

    ebayRows.push(fullHeader.join(","));

    // Helper function to escape CSV values
    const escapeCsv = (val) => {
      if (val === null || val === undefined) return "";
      const str = String(val);
      if (str.includes(",") || str.includes('"') || str.includes("\n")) {
        return `"${str.replace(/"/g, '""')}"`;
      }
      return str;
    };

    // Add data rows with item specifics
    ebayData.processedRows.slice(1).forEach((row, idx) => {
      const action = "Draft";
      const customSKU = row[31] || ""; // SKU column now includes /CostPerOne
      const categoryId = row[33] || "47155";
      const title = row[34] || "";
      const upc = String(row[7] || "");
      const price = row[32] || 0;
      const quantity = row[19] || 1;
      const photoUrl = [row[13], row[14], row[15], row[16], row[17], row[18]]
        .filter((img) => img && img !== "N/A" && img !== "")
        .join("|");
      const conditionId = 1500;

      // Use generated description
      let description = "";
      if (ebayData.allDescriptions && ebayData.allDescriptions[idx]) {
        description = String(ebayData.allDescriptions[idx] || "");
      }

      const dataBase = [
        action,
        escapeCsv(customSKU),
        categoryId,
        escapeCsv(title),
        upc,
        price,
        quantity,
        photoUrl,
        conditionId,
        escapeCsv(description),
        "FixedPrice",
      ];

      // Add specifics values for this item
      const dataSpecifics = ebayData.sortedSpecificsKeys.map((key) => {
        const value = ebayData.itemSpecificsArray[idx]?.[key] || "";
        return escapeCsv(value);
      });

      const fullRow = [...dataBase, ...dataSpecifics];
      ebayRows.push(fullRow.join(","));
    });

    const csvContent = ebayRows.join("\n");
    const csvBlob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const csvUrl = URL.createObjectURL(csvBlob);

    document.getElementById("downloadEbay").href = csvUrl;
    document.getElementById(
      "downloadEbay"
    ).download = `ebay_upload_${ebayData.dateSKU}.csv`;
    document.getElementById("downloadEbay").style.display = "inline-block";

    showStatus(
      `✅ Files ready! LISTING_${ebayData.dateSKU}.xlsx and ebay_upload_${ebayData.dateSKU}.csv`,
      "success"
    );

    displayPreview(
      ebayData.processedRows.slice(1, 11),
      ebayData.itemSpecificsArray,
      ebayData.allDescriptions
    );
  }

  // =============================================================================
  // PREVIEW TABLE DISPLAY
  // =============================================================================

  function displayPreview(rows, specificsArray, descriptionsArray) {
    const preview = document.getElementById("preview");
    const tableContainer = document.getElementById("previewTable");

    let html = `
    <table>
      <thead>
        <tr>
          <th>SKU</th>
          <th>Title</th>
          <th>Cost Per One</th>
          <th>Sales Price</th>
          <th>Category ID</th>
          <th>Type</th>
          <th>Key Specifics</th>
        </tr>
      </thead>
      <tbody>
  `;

    rows.forEach((row, rowIdx) => {
      const sku = row[31] || "";
      const title = row[34] || "";
      const costPerOne = row[30] || 0;
      const salesPrice = row[32] || 0;
      const categoryId = row[33] || "";

      // Get specifics for this row
      const specifics =
        specificsArray && specificsArray[rowIdx] ? specificsArray[rowIdx] : {};
      const itemType = specifics.Type || "—";

      // Get key specifics (non-empty ones, excluding Type)
      const keySpecs = [];
      for (const [key, value] of Object.entries(specifics)) {
        if (value && value !== "" && key !== "Type") {
          keySpecs.push(`${key}: ${value}`);
        }
      }
      const specsText =
        keySpecs.length > 0 ? keySpecs.slice(0, 3).join(" | ") : "—";

      html += `
      <tr>
        <td>${sku}</td>
        <td title="${title}">${title.substring(0, 30)}...</td>
        <td>£${parseFloat(costPerOne).toFixed(2)}</td>
        <td>£${parseFloat(salesPrice).toFixed(2)}</td>
        <td>${categoryId}</td>
        <td>${itemType}</td>
        <td>${specsText}</td>
      </tr>
    `;
    });

    html += `</tbody></table>`;
    tableContainer.innerHTML = html;
    preview.style.display = "block";
  }
}); // END OF DOMContentLoaded
