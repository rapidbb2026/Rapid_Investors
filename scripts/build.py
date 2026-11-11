"""Turn the Kobo export of the Investor Survey into the JSON the dashboard reads."""
import json, re, sys, datetime as dt
from collections import Counter
import openpyxl

BN = re.compile(r'[\u0980-\u09FF]+')
NOW = dt.datetime.now(dt.timezone.utc).replace(tzinfo=None) + dt.timedelta(hours=6)  # Dhaka time

SRC, GEO, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
wb = openpyxl.load_workbook(SRC)
ws = wb.worksheets[0]
rows = list(ws.iter_rows(values_only=True))
H, R = rows[0], [r for r in rows[1:] if any(v not in (None, '') for v in r)]

# The script was written against the 140-column export downloaded on 25 Oct 2026.
# Kobo's API export can add or reorder columns (notes, group headers, consent), so each column
# is found again by its header text and the rows are rebuilt in the original order.
ORIG_HEADERS = ["start", "end", "username", "Do you agree to take part in this survey? (আপনি কি এই জরিপে অংশগ্রহণে সম্মত?)", "Firm name (প্রতিষ্ঠানের নাম)", "Location (division)", "Location (district)", "Year operations began in Bangladesh (বাংলাদেশে কার্যক্রম শুরুর বছর)", "Zone status অঞ্চলের অবস্থা", "Respondent position (উত্তরদাতার পদ)", "May we re-contact you for clarification or quality assurance? কোনো বিষয় স্পষ্ট করার জন্য বা গুণমান নিশ্চিতকরণের লক্ষ্যে আমরা কি আপনার সাথে পুনরায় যোগাযোগ করতে পারি?", "Investment type by ownership (বিনিয়োগের ভিত্তিতে মালিকানার ধরন)", "Approximate foreign ownership share (বিদেশি মালিকানার আনুমানিক অংশ)", "Main sector (top level, one answer) প্রধান খাত", "Main product or service (sub-sector) (প্রধান পণ্য বা সেবা (উপ-খাত))", "Number of permanent workers at present (বর্তমানে স্থায়ী কর্মী সংখ্যা)", "Number of temporary workers at present (বর্তমানে অস্থায়ী কর্মী সংখ্যা)", "Share of sales exported directly (last completed fiscal year) (প্রত্যক্ষভাবে রপ্তানিকৃত বিক্রয়ের অংশ (সর্বশেষ সম্পূর্ণ অর্থবছর))", "Share of sales as deemed / indirect exports (supplies to export-oriented firms, e.g. under back-to-back LC) (পরোক্ষ রপ্তানি হিসেবে বিক্রয়ের অংশ (রপ্তানিমুখী প্রতিষ্ঠানকে সরবরাহ, যেমন ব্যাক-টু-ব্যাক LC-এর আওতায়))", "Value of fixed assets, excluding land and factory building (স্থায়ী সম্পদের মূল্য, জমি ও কারখানা ভবন বাদে)", "Annual turnover, last completed fiscal year (FY2025-26) (বার্ষিক টার্নওভার, সর্বশেষ সম্পূর্ণ অর্থবছর (FY2025-26))", "Did the firm import goods, raw materials, intermediates or equipment in the last 12 months? (গত 12 মাসে প্রতিষ্ঠানটি কি পণ্য, কাঁচামাল, মধ্যবর্তী পণ্য বা যন্ত্রপাতি আমদানি করেছে?)", "Main port or station of entry for the firm's imports (প্রতিষ্ঠানের আমদানির প্রধান বন্দর বা স্টেশন)", "Share of input and equipment expenditure that is imported (last completed fiscal year, FY2025-26) (আমদানিকৃত উপকরণ ও যন্ত্রপাতি ব্যয়ের অংশ (সর্বশেষ সম্পূর্ণ অর্থবছর, FY2025-26)", "**Top Importing Product**", "**Product Name**", "**Import Value**", "**HS code**", "##### Top_Importing_Product_1", "<span style=\"display:none\">row-Product Name</span>", "<span style=\"display:none\">row-Import Value</span>", "<span style=\"display:none\">row-HS code</span>", "##### Top_Importing_Product_2", "<span style=\"display:none\">row_1-Product Name</span>", "<span style=\"display:none\">row_1-Import Value</span>", "<span style=\"display:none\">row_1-HS code</span>", "##### Top_Importing_Product_3", "<span style=\"display:none\">row_2-Product Name</span>", "<span style=\"display:none\">row_2-Import Value</span>", "<span style=\"display:none\">row_2-HS code</span>", "Use of bonded warehouse, drawback, exemption or zone schemes (বন্ডেড ওয়্যারহাউস, ড্রব্যাক, অব্যাহতি বা অঞ্চলভিত্তিক সুবিধা প্রকল্পের ব্যবহার)", "Approximately what share of the firm's total production cost is made up of duties and taxes paid at import? (প্রতিষ্ঠানের মোট উৎপাদন ব্যয়ের আনুমানিক কত অংশ আমদানির সময় পরিশোধিত শুল্ক ও করের সমন্বয়ে গঠিত?)", "Annual value of imports, last completed fiscal year (বার্ষিক আমদানির মূল্য, সর্বশেষ সম্পূর্ণ অর্থবছর)", "Number of import consignments (bills of entry) in the last 12 months (গত 12 মাসে আমদানি চালানের (বিল অব এন্ট্রি) সংখ্যা)", "Does Border-tax burden compared with three years ago increase? (তিন বছর আগের তুলনায় আমদানি-করের চাপ বেড়েছে?)", "If higher or lower: what was the main change? (বেশি বা কম হলে: প্রধান পরিবর্তন কী ছিল?)", "Total duties and taxes paid at import, last completed fiscal year (আমদানির সময় পরিশোধিত মোট শুল্ক ও কর, সর্বশেষ সম্পূর্ণ অর্থবছর)", "**Main export product**", "**product name**", "**HS code**", "**export value**", "##### main export product_1", "<span style=\"display:none\">row-product name</span>", "<span style=\"display:none\">row-HS code</span>", "<span style=\"display:none\">row-export value</span>", "##### main export product_2", "<span style=\"display:none\">row_1-product name</span>", "<span style=\"display:none\">row_1-HS code</span>", "<span style=\"display:none\">row_1-export value</span>", "In the last three years, has the firm been involved in trading goods, or applied for any customs, tariff, or investment approval? (গত তিন বছরে প্রতিষ্ঠানটি কি পণ্য আমদানি বা রপ্তানির সাথে জড়িত ছিল, অথবা কোনো কাস্টমস, শুল্ক বা বিনিয়োগ অনুমোদনের জন্য আবেদন করেছে?)", "How serious is each of the following for the firm at present? (Show card C. Tick one box per row.) C1 বর্তমানে প্রতিষ্ঠানের জন্য নিচের প্রতিটি বিষয় কতটা গুরুতর? (কার্ড C দেখান। প্রতি সারিতে একটি ঘরে টিক দিন।)", "Tariff on imported raw materials (আমদানিকৃত কাঁচামালের উপর শুল্ক)", "Tariff on imported intermediate inputs (parts, components, fabrics, chemicals) (আমদানিকৃত মধ্যবর্তী উপকরণের (যন্ত্রাংশ, উপাদান, কাপড়, রাসায়নিক) উপর শুল্ক)", "Tariff on capital machinery and equipment (মূলধনী যন্ত্রপাতি ও সরঞ্জামের উপর শুল্ক)", "Duty on your imported inputs is as high as, or higher than, duty on the finished goods you compete with (আপনার আমদানিকৃত উপকরণের উপর শুল্ক, আপনি যেসব তৈরি পণ্যের সাথে প্রতিযোগিতা করেন তার শুল্কের সমান বা তার চেয়ে বেশি)", "Supplementary Duty (SD) burden (সম্পূরক শুল্ক (SD)-এর বোঝা)", "Regulatory Duty (RD) burden (নিয়ন্ত্রণমূলক শুল্ক (RD)-এর বোঝা)", "Delay or difficulty adjusting VAT and AT and AIT paid at import (আমদানির সময় পরিশোধিত VAT ও AT সমন্বয়ে বিলম্ব বা অসুবিধা)", "HS classification uncertainty or dispute (HS শ্রেণিবিন্যাসে অনিশ্চয়তা বা বিরোধ)", "Customs clearance delay (কাস্টমস ছাড়করণে বিলম্ব)", "Frequent SRO / policy changes, or inadequate notice (ঘন ঘন SRO / নীতি পরিবর্তন, অথবা অপর্যাপ্ত পূর্ব-বিজ্ঞপ্তি)", "Complex bonded, drawback or exemption procedures (জটিল বন্ডেড, ড্রব্যাক বা অব্যাহতি পদ্ধতি)", "Minimum import value or tariff value set above the price actually paid (প্রকৃত পরিশোধিত মূল্যের চেয়ে বেশি নির্ধারিত ন্যূনতম আমদানি মূল্য বা শুল্ক মূল্য)", "Through which main mechanism does the largest problem affect the firm? (কোন প্রধান প্রক্রিয়ার মাধ্যমে বৃহত্তম সমস্যাটি প্রতিষ্ঠানকে প্রভাবিত করে?", "Typical routine customs clearance time at the main port of entry (C4a) (প্রধান প্রবেশ বন্দরে সাধারণ নিয়মিত কাস্টমস ছাড়করণের সময় (C4a))", "Number of HS classification or valuation disputes in the last 12 months (গত 12 মাসে HS শ্রেণিবিন্যাস বা মূল্যায়ন বিরোধের সংখ্যা)", "Ease of obtaining reliable advance information on tariffs and procedures (শুল্ক ও পদ্ধতি সম্পর্কে নির্ভরযোগ্য অগ্রিম তথ্য পাওয়ার সহজতা)", "Effectiveness of bonded access (বন্ডেড সুবিধা প্রাপ্তির কার্যকারিতা)", "Typical duty drawback or refund time (সাধারণ ড্রব্যাক বা শুল্ক ফেরতের সময়)", "Main obstacle to using these schemes (শুল্কমুক্ত আমদানি সুবিধা ব্যবহারের প্রধান বাধা)", "Duty or tax paid but not yet refunded (drawback or VAT) at present (বর্তমানে শুল্ক বা কর পরিশোধিত কিন্তু এখনও ফেরত পাওয়া যায়নি (ড্রব্যাক বা VAT))", "Amount held as bank guarantee or bond security with customs at present (বর্তমানে কাস্টমসের কাছে ব্যাংক গ্যারান্টি বা বন্ড জামানত হিসেবে রাখা অর্থের পরিমাণ)", "How serious is each of the following for the firm at present? ( বর্তমানে প্রতিষ্ঠানের জন্য নিচের প্রতিটি বিষয় কতটা গুরুতর?)", "Land and site access (জমি ও স্থান প্রাপ্তি)", "Power reliability and cost (বিদ্যুতের নির্ভরযোগ্যতা ও ব্যয়)", "Transport, logistics and ports (পরিবহন, লজিস্টিকস ও বন্দর)", "Finance, trade finance and foreign exchange (অর্থায়ন, বাণিজ্য অর্থায়ন ও বৈদেশিক মুদ্রা)", "Regulatory approvals and licences (নিয়ন্ত্রক অনুমোদন ও লাইসেন্স)", "Coordination between government agencies (সরকারি সংস্থাসমূহের মধ্যে সমন্বয়)", "Shortage of Skilled labour (দক্ষ শ্রমশক্তির অভাব)", "Total investment by the firm in the last three fiscal years (FY2023-24 to FY2025-26), excluding land (গত তিন অর্থবছরে (FY2023-24 থেকে FY2025-26) জমি বাদে প্রতিষ্ঠানের মোট বিনিয়োগ", "In the last three years, have tariff or customs issues caused the firm to do any of the following (Assume that the firm is heavily dependent on trade but investment was delayed due to complexity in customs)? গত তিন বছরে শুল্ক বা কাস্টমস সংক্রান্ত সমস্যার কারণে প্রতিষ্ঠানটি কি নিচের কোনোটি করেছে?", "In the last three years, have tariff or customs issues caused the firm to do any of the following (Assume that the firm is heavily dependent on trade but investment was delayed due to complexity in customs)? গত তিন বছরে শুল্ক বা কাস্টমস সংক্রান্ত সমস্যার কারণে প্রতিষ্ঠানটি কি নিচের কোনোটি করেছে?/Delay an investment বিনিয়োগে বিলম্ব করেছে", "In the last three years, have tariff or customs issues caused the firm to do any of the following (Assume that the firm is heavily dependent on trade but investment was delayed due to complexity in customs)? গত তিন বছরে শুল্ক বা কাস্টমস সংক্রান্ত সমস্যার কারণে প্রতিষ্ঠানটি কি নিচের কোনোটি করেছে?/Defer an investment বিনিয়োগ পিছিয়ে দিয়েছে", "In the last three years, have tariff or customs issues caused the firm to do any of the following (Assume that the firm is heavily dependent on trade but investment was delayed due to complexity in customs)? গত তিন বছরে শুল্ক বা কাস্টমস সংক্রান্ত সমস্যার কারণে প্রতিষ্ঠানটি কি নিচের কোনোটি করেছে?/Reduce the scale of a planned investment পরিকল্পিত বিনিয়োগের মাত্রা কমিয়েছে", "In the last three years, have tariff or customs issues caused the firm to do any of the following (Assume that the firm is heavily dependent on trade but investment was delayed due to complexity in customs)? গত তিন বছরে শুল্ক বা কাস্টমস সংক্রান্ত সমস্যার কারণে প্রতিষ্ঠানটি কি নিচের কোনোটি করেছে?/Cancel an investment বিনিয়োগ বাতিল করেছে", "In the last three years, have tariff or customs issues caused the firm to do any of the following (Assume that the firm is heavily dependent on trade but investment was delayed due to complexity in customs)? গত তিন বছরে শুল্ক বা কাস্টমস সংক্রান্ত সমস্যার কারণে প্রতিষ্ঠানটি কি নিচের কোনোটি করেছে?/Use a less-preferred sourcing option কম পছন্দের সংগ্রহ বিকল্প ব্যবহার করেছে", "In the last three years, have tariff or customs issues caused the firm to do any of the following (Assume that the firm is heavily dependent on trade but investment was delayed due to complexity in customs)? গত তিন বছরে শুল্ক বা কাস্টমস সংক্রান্ত সমস্যার কারণে প্রতিষ্ঠানটি কি নিচের কোনোটি করেছে?/Lose export orders or competitiveness রপ্তানি অর্ডার বা প্রতিযোগিতা সক্ষমতা হারিয়েছে", "Which THREE of the following reforms are most urgent for the firm? (Show card G. Select up to three; do not rate every item.) G1 (নিচের কোন তিনটি সংস্কার প্রতিষ্ঠানের জন্য সবচেয়ে জরুরি? (কার্ড G দেখান। সর্বোচ্চ তিনটি নির্বাচন করুন; প্রতিটি বিষয়ে রেটিং দেবেন না।)", "Which THREE of the following reforms are most urgent for the firm? (Show card G. Select up to three; do not rate every item.) G1 (নিচের কোন তিনটি সংস্কার প্রতিষ্ঠানের জন্য সবচেয়ে জরুরি? (কার্ড G দেখান। সর্বোচ্চ তিনটি নির্বাচন করুন; প্রতিটি বিষয়ে রেটিং দেবেন না।)/Reduce or rationalise tariffs on selected inputs নির্বাচিত উপকরণের উপর শুল্ক হ্রাস বা যৌক্তিকীকরণ", "Which THREE of the following reforms are most urgent for the firm? (Show card G. Select up to three; do not rate every item.) G1 (নিচের কোন তিনটি সংস্কার প্রতিষ্ঠানের জন্য সবচেয়ে জরুরি? (কার্ড G দেখান। সর্বোচ্চ তিনটি নির্বাচন করুন; প্রতিটি বিষয়ে রেটিং দেবেন না।)/2 Simplify tariff bands and reduce escalation শুল্ক স্তর সহজ করা এবং শুল্ক বৃদ্ধির ধাপ (escalation) কমানো", "Which THREE of the following reforms are most urgent for the firm? (Show card G. Select up to three; do not rate every item.) G1 (নিচের কোন তিনটি সংস্কার প্রতিষ্ঠানের জন্য সবচেয়ে জরুরি? (কার্ড G দেখান। সর্বোচ্চ তিনটি নির্বাচন করুন; প্রতিটি বিষয়ে রেটিং দেবেন না।)/3 Improve HS guidance and advance rulings HS নির্দেশনা ও অগ্রিম রুলিং উন্নত করা", "Which THREE of the following reforms are most urgent for the firm? (Show card G. Select up to three; do not rate every item.) G1 (নিচের কোন তিনটি সংস্কার প্রতিষ্ঠানের জন্য সবচেয়ে জরুরি? (কার্ড G দেখান। সর্বোচ্চ তিনটি নির্বাচন করুন; প্রতিটি বিষয়ে রেটিং দেবেন না।)/4 Improve valuation transparency and dispute resolution মূল্যায়নের স্বচ্ছতা ও বিরোধ নিষ্পত্তি উন্নত করা", "Which THREE of the following reforms are most urgent for the firm? (Show card G. Select up to three; do not rate every item.) G1 (নিচের কোন তিনটি সংস্কার প্রতিষ্ঠানের জন্য সবচেয়ে জরুরি? (কার্ড G দেখান। সর্বোচ্চ তিনটি নির্বাচন করুন; প্রতিটি বিষয়ে রেটিং দেবেন না।)/5 Reduce clearance time through process and digital reform প্রক্রিয়া ও ডিজিটাল সংস্কারের মাধ্যমে ছাড়করণের সময় কমানো", "Which THREE of the following reforms are most urgent for the firm? (Show card G. Select up to three; do not rate every item.) G1 (নিচের কোন তিনটি সংস্কার প্রতিষ্ঠানের জন্য সবচেয়ে জরুরি? (কার্ড G দেখান। সর্বোচ্চ তিনটি নির্বাচন করুন; প্রতিটি বিষয়ে রেটিং দেবেন না।)/6 Simplify bonded, drawback and refund procedures বন্ডেড, ড্রব্যাক ও ফেরত পদ্ধতি সহজ করা", "Which THREE of the following reforms are most urgent for the firm? (Show card G. Select up to three; do not rate every item.) G1 (নিচের কোন তিনটি সংস্কার প্রতিষ্ঠানের জন্য সবচেয়ে জরুরি? (কার্ড G দেখান। সর্বোচ্চ তিনটি নির্বাচন করুন; প্রতিটি বিষয়ে রেটিং দেবেন না।)/7 Improve advance notice and consultation on SRO changes SRO পরিবর্তনের বিষয়ে অগ্রিম বিজ্ঞপ্তি ও পরামর্শ উন্নত করা", "Which THREE of the following reforms are most urgent for the firm? (Show card G. Select up to three; do not rate every item.) G1 (নিচের কোন তিনটি সংস্কার প্রতিষ্ঠানের জন্য সবচেয়ে জরুরি? (কার্ড G দেখান। সর্বোচ্চ তিনটি নির্বাচন করুন; প্রতিটি বিষয়ে রেটিং দেবেন না।)/8 Improve NBR–MoC–BTTC–IB coordination NBR–MoC–BTTC–IB সমন্বয় উন্নত করা", "Which THREE of the following reforms are most urgent for the firm? (Show card G. Select up to three; do not rate every item.) G1 (নিচের কোন তিনটি সংস্কার প্রতিষ্ঠানের জন্য সবচেয়ে জরুরি? (কার্ড G দেখান। সর্বোচ্চ তিনটি নির্বাচন করুন; প্রতিটি বিষয়ে রেটিং দেবেন না।)/9 Strengthen adjustment support for respective sectors (respondents own sector) সংশ্লিষ্ট খাতসমূহের জন্য সমন্বয় সহায়তা জোরদার করা (উত্তরদাতার নিজস্ব খাত)", "Confidence that tariff and customs rules will be remain same over the next two years (আগামী দুই বছরে শুল্ক ও কাস্টমস নিয়ম একই থাকবে, এই বিষয়ে আস্থা)", "Effectiveness of government–business consultation on tariff changes (শুল্ক পরিবর্তন বিষয়ে সরকার–ব্যবসা পরামর্শের কার্যকারিতা)", "Overall tariff and customs environment for investment (বিনিয়োগের জন্য সামগ্রিক শুল্ক ও কাস্টমস পরিবেশ)", "Is there a specific tariff line, SRO or customs procedure that the study should examine? এমন কোনো নির্দিষ্ট শুল্ক লাইন, SRO বা কাস্টমস পদ্ধতি আছে কি, যা এই গবেষণায় পরীক্ষা করা উচিত?", "Record your current location", "_Record your current location_latitude", "_Record your current location_longitude", "_Record your current location_altitude", "_Record your current location_precision", "S2. Is the enterprise currently operating in Bangladesh? প্রতিষ্ঠানটি কি বর্তমানে বাংলাদেশে কার্যরত?", "S3. Location (division)", "S4. Zone status অঞ্চলের অবস্থা", "S6. Who handles customs clearance of the firm's imports and exports? প্রতিষ্ঠানটির আমদানি ও রপ্তানির কাস্টমস ক্লিয়ারেন্স (শুল্ক খালাস) কে পরিচালনা করে?", "A7. Primary investment type (প্রাথমিক বিনিয়োগের ধরন) Confusing Question", "B3. Main imported category by value (মূল্যের ভিত্তিতে প্রধান আমদানিকৃত শ্রেণি)", "C1 How serious is each of the following for the firm at present? (Show card C. Tick one box per row.) C1 বর্তমানে প্রতিষ্ঠানের জন্য নিচের প্রতিটি বিষয় কতটা গুরুতর? (কার্ড C দেখান। প্রতি সারিতে একটি ঘরে টিক দিন।)", "C1g Customs valuation uncertainty or dispute কাস্টমস মূল্যায়নে অনিশ্চয়তা বা বিরোধ", "C1i Inconsistent or discretionary interpretation অসামঞ্জস্যপূর্ণ বা স্বেচ্ছাধীন ব্যাখ্যা", "Ask if the firm uses or is eligible for any scheme in B4 (code 1 or 2). প্রতিষ্ঠান B4-এর কোনো প্রকল্প (কোড 1 বা 2) ব্যবহার করলে বা তার জন্য যোগ্য হলে জিজ্ঞাসা করুন।", "C1m Competition from goods imported informally or under-invoiced অনানুষ্ঠানিকভাবে বা কম-চালানকৃত (under-invoiced) পণ্য আমদানির সাথে প্রতিযোগিতা", "E1 How serious is each of the following for the firm at present? E1 বর্তমানে প্রতিষ্ঠানের জন্য নিচের প্রতিটি বিষয় কতটা গুরুতর?", "F2 If any investment was delayed, reduced or cancelled (F1 a–d): approximate value affected কোনো বিনিয়োগ বিলম্বিত, হ্রাস বা বাতিল হয়ে থাকলে (F1 a–d): প্রভাবিত আনুমানিক মূল্য", "F3 Approximate number of jobs not created or lost as a result এর ফলে সৃষ্টি না হওয়া বা হারানো চাকরির আনুমানিক সংখ্যা", "_id", "_uuid", "_submission_time", "_validation_status", "_status", "_submitted_by", "__version__", "meta/rootUuid", "_index"]

def _norm(x):
    return re.sub(r'\s+', ' ', str(x or '')).strip()

def _loose(x):
    """Header without Bangla, option numbers or punctuation, for headers whose wording was edited slightly."""
    x = BN.sub(' ', _norm(x)).lower()
    x = re.sub(r'/\s*\d+\s*(=\s*)?', '/', x)
    return re.sub(r'[^a-z0-9/]+', '', x)

def _remap(header, rows):
    def build(key):
        where = {}
        for j, t in enumerate(header):
            where.setdefault(key(t), []).append(j)
        seen, out = {}, []
        for t in ORIG_HEADERS:
            kt = key(t); k = seen.get(kt, 0); seen[kt] = k + 1
            hits = where.get(kt, [])
            out.append(hits[k] if k < len(hits) else None)
        return out
    idx, loose = build(_norm), build(_loose)
    idx = [a if a is not None else b for a, b in zip(idx, loose)]
    # last resort: a column sitting between two found neighbours that are one column apart
    for i in range(1, len(idx) - 1):
        if idx[i] is None and idx[i - 1] is not None and idx[i + 1] is not None and idx[i + 1] - idx[i - 1] == 2:
            idx[i] = idx[i - 1] + 1
            print('Matched by position (heading wording changed): %s -> %s' % (ORIG_HEADERS[i][-70:], _norm(header[idx[i]])[-70:]))
    return idx

IDX = _remap(H, R)
USED = sorted(set([0, 1, 3, 4, 6, 7, 12, 14, 113, 114, 116, 131, 133, 136, 137, 139]
                  + list(range(5, 24)) + [29, 30, 33, 34, 37, 38] + list(range(40, 47))
                  + [52, 54, 56, 58, 59] + list(range(61, 91)) + list(range(92, 98)) + list(range(99, 111))))
missing = [i for i in USED if IDX[i] is None]
if missing:
    sys.exit('These questions are missing from the Kobo export, so build.py needs updating:\n' +
             '\n'.join('  %d: %s' % (i, ORIG_HEADERS[i][:120]) for i in missing))
R = [tuple(r[j] if j is not None and j < len(r) else None for j in IDX) for r in R]
H = ORIG_HEADERS

def when(v):
    """Kobo gives a datetime in a downloaded file but text like 2026-10-19T10:44:34.535+06:00 from the API."""
    if v in (None, ''): return None
    if isinstance(v, dt.datetime): x = v
    else:
        x = dt.datetime.fromisoformat(str(v).strip().replace('Z', '+00:00').replace(' ', 'T', 1))
    if x.tzinfo is not None:
        x = x.astimezone(dt.timezone(dt.timedelta(hours=6))).replace(tzinfo=None)
    return x

OWN_FIX = {'Foreign / MNC ( / (MNC))': 'Foreign / MNC', 'MNC ( )': 'MNC'}
BN = re.compile(r'[ঀ-৿]+')
def clean(v):
    if v is None: return None
    s = str(v)
    s = BN.sub('', s)
    s = re.sub(r'^\s*\d+\s*=\s*', '', s)
    s = s.replace(';', '').replace('\n', ' ')
    s = re.sub(r'\s+', ' ', s).strip()
    s = OWN_FIX.get(s, s)
    for _ in range(3):
        s = re.sub(r'\([\s/,\-]*\)', '', s).strip()
        m = re.search(r'\(([^()]*)\)\s*$', s)
        if m and m.group(1).strip() and m.group(1).strip().lower() in s[:m.start()].lower():
            s = s[:m.start()].strip()
    s = re.sub(r'\s+', ' ', s).strip(' ,')
    return s or None

def num(v):
    if v in (None, ''): return None
    try: return float(str(v).replace(',', '').strip())
    except ValueError: return None

def money(v):
    """Amount from free text. Returns (value or None, 'BDT' or 'USD', problem or None).
    Handles '25 crore bdt', '1.5 lakh', '500000', 'Us 163.39k', '$2m'."""
    if v in (None, ''): return None, 'BDT', None
    s = str(v).strip().lower().replace(',', '')
    cur = 'USD' if re.search(r'(^|[^a-z])(us|usd)([^a-z]|$)|\$|dollar', s) else 'BDT'
    s = re.sub(r'usd|us|\$|dollars?|bdt|tk|taka', ' ', s).strip(' -:=')
    if re.fullmatch(r'\d+\.\d{3}', s): return None, cur, 'Value is unclear (the dot could be a decimal point or a thousands separator)'
    m = re.match(r'^([\d.]+)\s*(k|thousand|m|mn|million|crore|core|cr|lakh|lac)?\.?$', s)
    if not m: return None, cur, 'Value is not a number'
    try: x = float(m.group(1))
    except ValueError: return None, cur, 'Value is not a number'
    x *= {'k': 1e3, 'thousand': 1e3, 'm': 1e6, 'mn': 1e6, 'million': 1e6, 'crore': 1e7, 'core': 1e7,
          'cr': 1e7, 'lakh': 1e5, 'lac': 1e5, None: 1}[m.group(2)]
    if x == 0: return None, cur, 'Value is zero'
    return x, cur, None

# ---------- question schema ----------
SEV = ['none', 'minor', 'moderate', 'serious', 'very serious']
Q = []  # each: key, col, section, label, type, suffix
def add(key, col, sec, label, typ='single', **kw):
    Q.append(dict(key=key, col=col, section=sec, label=label, type=typ, **kw))

S, A, B, C, D, E, F, G, HH = ('Screener', 'A. Firm and respondent profile', 'B. Trade and tariff exposure',
    'C. Tariff and customs bottlenecks', 'D. Bonded warehouse, drawback and facilitation',
    'E. Wider investment-climate constraints', 'F. Business and investment impacts', 'G. Reform priorities',
    'H. Predictability and institutional experience')
YN = ['yes', 'no']
# Screener
add('div', 5, S, 'Location (division)',
    order=['Barishal', 'Chattogram', 'Dhaka', 'Khulna', 'Mymensingh', 'Rajshahi', 'Rangpur', 'Sylhet'])
add('year', 7, S, 'Year operations began in Bangladesh', 'num', unit='year')
add('zone', 8, S, 'Zone status', order=['EPZ (BEPZA)', 'Economic Zone (BEZA)', 'Hi-Tech Park (BHTPA)', 'outside zones'])
add('pos', 9, S, 'Respondent position',
    order=['Owner / director', 'CEO / GM', 'Finance / tax', 'Customs / trade', 'Supply chain', 'Strategy', 'Other'])
add('recontact', 10, S, 'May we re-contact you for clarification or quality assurance?', order=YN)
# Module A
add('own', 11, A, 'Investment type by ownership',
    order=['Domestic private', 'Foreign', 'MNC', 'Foreign / MNC', 'Joint venture', 'State-owned', 'Other'])
add('own_share', 12, A, 'Approximate foreign ownership share', ordinal=True)
add('sector', 13, A, 'Main sector (top level, one answer)',
    order=['Manufacturing (incl. agro-processing and pharmaceuticals)', 'Agriculture, fisheries and livestock',
           'Services (incl. ICT and logistics)'])
add('perm', 15, A, 'Number of permanent workers at present', 'num', unit='workers')
add('temp', 16, A, 'Number of temporary workers at present', 'num', unit='workers')
add('assets', 19, A, 'Value of fixed assets, excluding land and factory building', ordinal=True)
add('turnover', 20, A, 'Annual turnover, last completed fiscal year (FY2025-26)', ordinal=True)
add('trade_app', 59, A, 'In the last three years, has the firm traded goods, or applied for any customs, tariff or investment approval?', order=YN)
add('exp_dir', 17, A, 'Share of sales exported directly (last completed fiscal year)', ordinal=True)
add('exp_ind', 18, A, 'Share of sales as deemed / indirect exports', ordinal=True)
add('scheme', 40, A, 'Use of bonded warehouse, drawback, exemption or zone schemes',
    order=['uses', 'eligible but does not use', 'not eligible', 'unaware'])
# Module B
add('imp', 21, B, 'Did the firm import goods, raw materials, intermediates or equipment in the last 12 months?', order=YN)
add('port', 22, B, "Main port or station of entry for the firm's imports",
    order=['Chattogram port', 'Dhaka ICD (Kamalapur / Pangaon)', 'Dhaka airport (HSIA)', 'Benapole', 'Mongla', 'Other'])
add('imp_share', 23, B, 'Share of input and equipment expenditure that is imported (FY2025-26)', ordinal=True, suffix='%')
add('duty_cost', 41, B, "Share of the firm's total production cost made up of duties and taxes paid at import", ordinal=True)
add('imp_value', 42, B, 'Annual value of imports, last completed fiscal year', ordinal=True)
add('consign', 43, B, 'Number of import consignments (bills of entry) in the last 12 months', ordinal=True)
add('burden', 44, B, 'Border-tax burden compared with three years ago',
    order=['Much lower', 'Lower', 'Same', 'Higher', 'Much higher'])
add('burden_change', 45, B, 'If higher or lower: what was the main change?',
    order=['New or higher RD / SD', 'Change in CD rate', 'Higher assessed (tariff) value', 'Exchange rate',
           'Loss of exemption or concession', 'Change in products imported', 'Other'])
add('duty_paid', 46, B, 'Total duties and taxes paid at import, last completed fiscal year', ordinal=True)
# Module C
C1 = [(61, 'Tariff on imported raw materials'),
      (62, 'Tariff on imported intermediate inputs (parts, components, fabrics, chemicals)'),
      (63, 'Tariff on capital machinery and equipment'),
      (64, 'Duty on imported inputs as high as, or higher than, duty on competing finished goods'),
      (65, 'Supplementary Duty (SD) burden'), (66, 'Regulatory Duty (RD) burden'),
      (67, 'Delay or difficulty adjusting VAT, AT and AIT paid at import'),
      (68, 'HS classification uncertainty or dispute'), (69, 'Customs clearance delay'),
      (70, 'Frequent SRO / policy changes, or inadequate notice'),
      (71, 'Complex bonded, drawback or exemption procedures'),
      (72, 'Minimum import value or tariff value set above the price actually paid')]
for i, (c, l) in enumerate(C1):
    add('c1_%d' % (i + 1), c, C, l, 'sev', order=SEV, grid='c1')
add('mech', 73, C, 'Main mechanism through which the largest problem affects the firm',
    order=['Input cost', 'Cash flow', 'Delay', 'Uncertainty', 'Compliance cost', 'Lost order', 'Investment delay', 'Other'])
add('clear_time', 74, C, 'Typical routine customs clearance time at the main port of entry', ordinal=True, suffix=' days')
add('disputes', 75, C, 'Number of HS classification or valuation disputes in the last 12 months', ordinal=True)
add('info', 76, C, 'Ease of obtaining reliable advance information on tariffs and procedures',
    order=['Very easy', 'Easy', 'Difficult', 'Very difficult'])
# Module D
add('bonded_eff', 77, D, 'Effectiveness of bonded access',
    order=['Very effective', 'effective', 'neutral', 'ineffective', 'very ineffective', 'not applicable (N/A)'])
add('drawback_time', 78, D, 'Typical duty drawback or refund time', ordinal=True, suffix=' months')
add('obstacle', 79, D, 'Main obstacle to using these schemes',
    order=['Eligibility', 'Documentation', 'Delay', 'Uncertainty', 'Audit risk', 'Lack of information', 'Agent cost', 'Other'])
add('unrefunded', 80, D, 'Duty or tax paid but not yet refunded (drawback or VAT) at present', ordinal=True)
add('guarantee', 81, D, 'Amount held as bank guarantee or bond security with customs at present', ordinal=True)
# Module E
E1 = [(83, 'Land and site access'), (84, 'Power reliability and cost'), (85, 'Transport, logistics and ports'),
      (86, 'Finance, trade finance and foreign exchange'), (87, 'Regulatory approvals and licences'),
      (88, 'Coordination between government agencies'), (89, 'Shortage of skilled labour')]
for i, (c, l) in enumerate(E1):
    add('e1_%d' % (i + 1), c, E, l, 'sev', order=SEV, grid='e1')
# Module F
add('invest3', 90, F, 'Total investment in the last three fiscal years (FY2023-24 to FY2025-26), excluding land', ordinal=True)
F1 = [(92, 'Delay an investment'), (93, 'Defer an investment'), (94, 'Reduce the scale of a planned investment'),
      (95, 'Cancel an investment'), (96, 'Use a less-preferred sourcing option'), (97, 'Lose export orders or competitiveness')]
add('f1', None, F, 'In the last three years, have tariff or customs issues caused the firm to do any of the following?',
    'multi', cols=F1, none_label='None selected')
# Module G
G1 = [(99, 'Reduce or rationalise tariffs on selected inputs'), (100, 'Simplify tariff bands and reduce escalation'),
      (101, 'Improve HS guidance and advance rulings'), (102, 'Improve valuation transparency and dispute resolution'),
      (103, 'Reduce clearance time through process and digital reform'),
      (104, 'Simplify bonded, drawback and refund procedures'),
      (105, 'Improve advance notice and consultation on SRO changes'), (106, 'Improve NBR–MoC–BTTC–IB coordination'),
      (107, "Strengthen adjustment support for the respondent's own sector")]
add('g1', None, G, 'Which THREE of the following reforms are most urgent for the firm?', 'multi', cols=G1,
    none_label='None selected')
# Module H
add('confidence', 108, HH, 'Confidence that tariff and customs rules will remain the same over the next two years',
    order=['Very confident', 'confident', 'moderately confident', 'less confident', 'not at all confident', 'DK'])
add('consult', 109, HH, 'Effectiveness of government–business consultation on tariff changes',
    order=['very effective', 'effective', 'moderately effective', 'ineffective', 'very ineffective', 'no experience'])
add('overall', 110, HH, 'Overall tariff and customs environment for investment',
    order=['very supportive', 'supportive', 'neutral', 'Constraining', 'Very constraining'])

LAST = re.compile(r"^(dk|don't know|dk / refused|varies|not applicable \(n/a\)|no experience)$", re.I)
def code_of(raw):
    m = re.match(r'^\s*(\d+)\s*=', str(raw)) if raw is not None else None
    return int(m.group(1)) if m else None

def label_for(q, raw):
    s = clean(raw)
    if s is None: return None
    if q['type'] == 'sev': s = s.lower()
    suf = q.get('suffix')
    if suf and re.search(r'\d$', s) and not LAST.match(s): s = s + suf
    for o in q.get('order') or []:
        if o.lower() == s.lower(): return o
    return s

# ---------- geography ----------
geo = json.load(open(GEO))
def pip(x, y, ring):
    ins, j = False, len(ring) - 1
    for i in range(len(ring)):
        xi, yi = ring[i]; xj, yj = ring[j]
        if (yi > y) != (yj > y) and x < (xj - xi) * (y - yi) / (yj - yi + 1e-12) + xi: ins = not ins
        j = i
    return ins
def which_div(lon, lat):
    for f in geo['features']:
        g = f['geometry']; polys = g['coordinates'] if g['type'] == 'MultiPolygon' else [g['coordinates']]
        for p in polys:
            if pip(lon, lat, p[0]): return f['properties']['name']
    return None

DHAKA_PLACES = re.compile(r'dhanmondi|motijheel|gulistan|chankharpool|paltan|polton|^dh\d*$', re.I)
UPAZILA = {'mirsharai': 'Chattogram'}

# ---------- records ----------
recs, issues = [], []
opt_codes = {q['key']: {} for q in Q}
name_seen = {}
def issue(rec, text, kind='bad', typ=''):
    issues.append(dict(id=rec['id'], firm=rec['firm'], issue=text, kind=kind, type=typ, enum=rec['enum'], date=rec['date']))

not_consented = 0
for r in R:
    if str(r[3] or '').strip().lower() == 'no':
        not_consented += 1
        continue
    start, end = when(r[0]), when(r[1])
    rec = dict(id=str(r[131]), idx=int(float(r[139])) if r[139] not in (None, '') else None,
               firm=re.sub(r'\s+', ' ', str(r[4] or '')).strip(),
               date=start.strftime('%Y-%m-%d') if start else None,
               start=start.strftime('%Y-%m-%d %H:%M') if start else None,
               mins=round((end - start).total_seconds() / 60, 1) if start and end else None,
               enum=(r[136] or '').strip() or 'Not recorded',
               version=str(r[137] or ''), product=clean(r[14]) or '',
               a={})
    for q in Q:
        if q['type'] == 'multi':
            picks = [lab for c, lab in q['cols'] if str(r[c]).strip() == '1']
            rec['a'][q['key']] = picks if picks else [q['none_label']]
        elif q['type'] == 'num':
            rec['a'][q['key']] = num(r[q['col']])
        else:
            raw = r[q['col']]
            lab = label_for(q, raw)
            rec['a'][q['key']] = lab
            if lab is not None and code_of(raw) is not None:
                opt_codes[q['key']][lab] = code_of(raw)
    # district as entered vs cleaned
    d_raw = re.sub(r'\s+', ' ', str(r[6] or '')).strip()
    d = d_raw
    if DHAKA_PLACES.search(d_raw): d = 'Dhaka'
    elif d_raw.lower() in UPAZILA: d = UPAZILA[d_raw.lower()]
    rec['dist'] = d; rec['dist_raw'] = d_raw
    # group
    own = rec['a']['own'] or ''
    rec['grp'] = 'Foreign or JV' if re.search(r'foreign|mnc|joint', own, re.I) else 'Domestic'
    rec['inzone'] = (rec['a']['zone'] or '').lower() != 'outside zones'
    rec['a']['age'] = (NOW.year - rec['a']['year']) if rec['a']['year'] else None
    rec['a']['workers'] = (rec['a']['perm'] or 0) + (rec['a']['temp'] or 0) if rec['a']['perm'] is not None else None
    # products (name, value; the value may be written in Tk or US$)
    def prod(pc, vc, what):
        if r[pc] in (None, ''): return None
        v, cur, prob = money(r[vc])
        raw = str(r[vc]).strip() if r[vc] is not None else ''
        if prob: issue(rec, '%s value for "%s": %s (%s)' % (what, clean(r[pc]), prob.lower(), raw), 'warn', 'Product value format')
        return dict(name=clean(r[pc]), value=v, cur=cur, raw=raw)
    imps = [p for p in (prod(29, 30, 'Import'), prod(33, 34, 'Import'), prod(37, 38, 'Import')) if p]
    exps = [p for p in (prod(52, 54, 'Export'), prod(56, 58, 'Export')) if p]
    rec['imports'], rec['exports'] = imps, exps
    lat, lon, prec = num(r[113]), num(r[114]), num(r[116])
    rec['gps'] = [lon, lat] if lat and lon else None
    rec['gps_prec'] = prec

    # ---- data checks ----
    if rec['mins'] is not None and rec['mins'] < 10:
        issue(rec, 'Interview took %s minutes (under 10)' % rec['mins'], 'bad', 'Very short interview')
    if rec['mins'] is not None and rec['mins'] > 240:
        issue(rec, 'Form was open for %s hours' % round(rec['mins'] / 60, 1), 'warn', 'Form left open')
    if d != d_raw:
        issue(rec, 'District entered as "%s" (counted as %s)' % (d_raw, d), 'warn', 'District entry')
    if rec['gps']:
        gd = which_div(lon, lat)
        if prec == 0: issue(rec, 'GPS precision is 0, so the location was probably typed in', 'warn', 'GPS')
        if gd is None: issue(rec, 'GPS point falls outside Bangladesh', 'bad', 'GPS')
        elif rec['a']['div'] and gd != rec['a']['div']:
            issue(rec, 'GPS point is in %s but the form says %s' % (gd, rec['a']['div']), 'bad', 'GPS')
    else:
        issue(rec, 'No GPS location', 'bad', 'GPS')
    if rec['a']['imp'] == 'no' and imps:
        issue(rec, 'Says no imports in 12 months but lists imported products', 'bad', 'Contradiction')
    if rec['a']['exp_dir'] == '0%' and rec['a']['exp_ind'] == '0%' and exps:
        issue(rec, 'Reports 0% exports but lists export products', 'warn', 'Contradiction')
    if rec['a']['perm'] is not None and rec['a']['perm'] == 0:
        issue(rec, 'Zero permanent workers', 'warn', 'Workers')
    key = re.sub(r'[^a-z]', '', rec['firm'].lower())
    if key in name_seen:
        issue(rec, 'Same firm name as record %s, possible duplicate' % name_seen[key], 'bad', 'Possible duplicate')
    else:
        name_seen[key] = rec['id']
    recs.append(rec)

# follow-up asked of everyone whose border-tax burden went up or down, but blank in every record
moved = [rc for rc in recs if rc['a']['burden'] in ('Much lower', 'Lower', 'Higher', 'Much higher')]
if moved and not any(rc['a']['burden_change'] for rc in recs):
    issues.append(dict(id='', firm='All records', enum='', date='', kind='warn', type='Skipped follow-up',
        issue='%d firms said the border-tax burden went up or down, but "what was the main change?" is blank for all of them. Check the skip logic in the form.' % len(moved)))
if not any(rc['a']['own_share'] for rc in recs):
    issues.append(dict(id='', firm='All records', enum='', date='', kind='warn', type='Skipped follow-up',
        issue='"Approximate foreign ownership share" is blank for every record, including all foreign, MNC and joint-venture firms. Check the skip logic in the form.'))

# form-version label conflicts
for q in Q:
    if q.get('ordinal'):
        by_code = {}
        for lab, c in opt_codes[q['key']].items(): by_code.setdefault(c, set()).add(lab)
        clash = {c: sorted(v) for c, v in by_code.items() if len(v) > 1}
        if clash:
            for c, labs in clash.items():
                issues.append(dict(id='', firm='All records', enum='', date='', kind='warn', type='Form version wording',
                    issue='"%s": answer code %d has different wording across form versions (%s)' % (q['label'], c, ' / '.join(labs))))

# option order
def order_opts(q):
    vals = Counter()
    for rc in recs:
        v = rc['a'][q['key']]
        for x in (v if isinstance(v, list) else [v]):
            if x is not None: vals[x] += 1
    if q['type'] == 'multi':
        labs = [l for _, l in q['cols']] + [q['none_label']]
        return labs
    if q.get('order'):
        o = list(q['order']); low = [x.lower() for x in o]
        extra = [k for k in vals if k.lower() not in low]
        return o + sorted(extra, key=lambda k: -vals[k])
    if q.get('ordinal'):
        codes = opt_codes[q['key']]
        def k(lab):
            if lab in codes: return (1, codes[lab], lab)
            if lab.lower() == 'none': return (0, 0, lab)
            if LAST.match(lab): return (3, 0, lab)
            return (2, -vals[lab], lab)
        return sorted(vals, key=k)
    return sorted(vals, key=lambda k: (bool(LAST.match(k)), -vals[k], k))

schema = []
for q in Q:
    item = dict(key=q['key'], section=q['section'], label=q['label'], type=q['type'])
    if q['type'] != 'num': item['options'] = order_opts(q)
    if q.get('unit'): item['unit'] = q['unit']
    if q.get('grid'): item['grid'] = q['grid']
    schema.append(item)
NUMS = [('year', S, 'Year operations began in Bangladesh', 'year'), ('age', S, 'Years operating in Bangladesh', 'years'),
        ('perm', A, 'Number of permanent workers at present', 'workers'),
        ('temp', A, 'Number of temporary workers at present', 'workers'),
        ('workers', A, 'All workers (permanent + temporary)', 'workers'),
        ('imp_v1', B, 'Import value, top importing product 1', 'money'), ('imp_v2', B, 'Import value, top importing product 2', 'money'),
        ('imp_v3', B, 'Import value, top importing product 3', 'money'),
        ('exp_v1', B, 'Export value, main export product 1', 'money'), ('exp_v2', B, 'Export value, main export product 2', 'money')]
schema = [s for s in schema if s['type'] != 'num'] + [dict(key=k, section=s, label=l, type='num', unit=u) for k, s, l, u in NUMS]

versions = Counter(rc['version'] for rc in recs)
last_sub = max(str(r[133]) for r in R if r[133])
out = dict(title='Investor Survey 2026', updated=max(rc['start'] for rc in recs),
           last_submission_utc=last_sub, exported=(sys.argv[4] if len(sys.argv) > 4 else NOW.strftime('%Y-%m-%d %H:%M')),
           schema=schema, records=recs, issues=issues, not_consented=not_consented,
           versions=[dict(v=v, n=n, first=min(rc['date'] for rc in recs if rc['version'] == v),
                          last=max(rc['date'] for rc in recs if rc['version'] == v)) for v, n in versions.most_common()])
json.dump(out, open(OUT, 'w'), ensure_ascii=False, separators=(',', ':'))
print('records', len(recs), 'issues', len(issues))
for i in issues: print(' -', i['id'], i['issue'])
print(Counter(rc['grp'] for rc in recs))
