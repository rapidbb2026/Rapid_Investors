"""Turn the Kobo export of the Investor Survey into the JSON the dashboard reads."""
import json, re, sys, datetime as dt
from collections import Counter
import openpyxl

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

def _remap(header, rows):
    where = {}
    for j, t in enumerate(header):
        where.setdefault(_norm(t), []).append(j)
    seen, idx = {}, []
    for t in ORIG_HEADERS:
        k = seen.get(t, 0); seen[t] = k + 1
        hits = where.get(t, [])
        idx.append(hits[k] if k < len(hits) else None)
    return idx

IDX = _remap(H, R)
USED = sorted(set([0, 1, 4, 6, 7, 12, 14, 111, 113, 114, 116, 131, 133, 136, 137, 139]
                  + list(range(5, 24)) + [29, 30, 31, 33, 34, 35, 37, 38, 39] + list(range(40, 45)) + [46]
                  + [52, 53, 54, 56, 57, 58, 59] + list(range(61, 91)) + list(range(92, 98)) + list(range(99, 112))))
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
    """Tk value from free text. Returns (value or None, problem or None)."""
    if v in (None, ''): return None, None
    s = str(v).strip().lower().replace(',', '')
    if re.search(r'us|\$', s): return None, 'Value given in US dollars'
    m = re.match(r'^([\d.]+)\s*(crore|cr|lakh|lac)?\s*(bdt|tk|taka)?$', s)
    if not m: return None, 'Value is not a number'
    x = float(m.group(1)) * {'crore': 1e7, 'cr': 1e7, 'lakh': 1e5, 'lac': 1e5, None: 1}[m.group(2)]
    if x == 0: return None, 'Value is zero'
    return x, None

# ---------- question schema ----------
SEV = ['none', 'minor', 'moderate', 'serious', 'very serious']
Q = []  # each: key, col, section, label, type, suffix
def add(key, col, sec, label, typ='single', **kw):
    Q.append(dict(key=key, col=col, section=sec, label=label, type=typ, **kw))

A, B, C, E, F, G = ('A. Firm profile', 'B. Trade and imports', 'C. Tariff and customs',
                    'E. Investment climate', 'F. Effects on investment', 'G. Reform priorities and outlook')
add('div', 5, A, 'Location (division)')
add('zone', 8, A, 'Zone status')
add('pos', 9, A, 'Respondent position')
add('recontact', 10, A, 'Agreed to be re-contacted')
add('own', 11, A, 'Investment type by ownership')
add('sector', 13, A, 'Main sector')
add('year', 7, A, 'Year operations began in Bangladesh', 'num', unit='year')
add('perm', 15, A, 'Permanent workers at present', 'num', unit='workers')
add('temp', 16, A, 'Temporary workers at present', 'num', unit='workers')
add('exp_dir', 17, A, 'Share of sales exported directly (last FY)', ordinal=True)
add('exp_ind', 18, A, 'Share of sales as deemed or indirect exports', ordinal=True)
add('assets', 19, A, 'Value of fixed assets, excluding land and building', ordinal=True)
add('turnover', 20, A, 'Annual turnover, FY2025-26', ordinal=True)
add('imp', 21, B, 'Imported goods or equipment in the last 12 months')
add('port', 22, B, 'Main port or station of entry for imports')
add('imp_share', 23, B, 'Share of input and equipment spending that is imported', ordinal=True, suffix='%')
add('scheme', 40, B, 'Use of bonded warehouse, drawback, exemption or zone schemes')
add('duty_cost', 41, B, 'Import duties and taxes as a share of production cost', ordinal=True)
add('imp_value', 42, B, 'Annual value of imports, last FY', ordinal=True)
add('consign', 43, B, 'Import consignments (bills of entry), last 12 months', ordinal=True)
add('burden', 44, B, 'Border-tax burden compared with three years ago',
    order=['Much lower', 'Lower', 'Same', 'Higher', 'Much higher'])
add('duty_paid', 46, B, 'Total duties and taxes paid at import, last FY', ordinal=True)
add('trade_app', 59, B, 'Traded or applied for a customs, tariff or investment process in the last 3 years')
C1 = [(61, 'Tariff on imported raw materials'), (62, 'Tariff on intermediate inputs'),
      (63, 'Tariff on capital machinery'), (64, 'Input duty as high as duty on competing finished goods'),
      (65, 'Supplementary Duty (SD) burden'), (66, 'Regulatory Duty (RD) burden'),
      (67, 'Delay adjusting VAT, AT and AIT paid at import'), (68, 'HS classification uncertainty or dispute'),
      (69, 'Customs clearance delay'), (70, 'Frequent SRO or policy changes'),
      (71, 'Complex bonded, drawback or exemption procedures'), (72, 'Minimum import or tariff value above price paid')]
for i, (c, l) in enumerate(C1):
    add('c1_%d' % (i + 1), c, C, l, 'sev', order=SEV, grid='c1')
add('mech', 73, C, 'Main way the largest problem affects the firm')
add('clear_time', 74, C, 'Routine customs clearance time at the main port', ordinal=True, suffix=' days')
add('disputes', 75, C, 'HS classification or valuation disputes, last 12 months (number)', ordinal=True)
add('info', 76, C, 'Ease of getting reliable advance information on tariffs')
add('bonded_eff', 77, C, 'Effectiveness of bonded access',
    order=['very ineffective', 'ineffective', 'neutral', 'effective', 'Very effective', 'not applicable (N/A)'])
add('drawback_time', 78, C, 'Typical duty drawback or refund time', ordinal=True, suffix=' months')
add('obstacle', 79, C, 'Main obstacle to using duty-free import schemes')
add('unrefunded', 80, C, 'Duty or tax paid but not yet refunded', ordinal=True)
add('guarantee', 81, C, 'Amount held as bank guarantee or bond with customs', ordinal=True)
E1 = [(83, 'Land and site access'), (84, 'Power reliability and cost'), (85, 'Transport, logistics and ports'),
      (86, 'Finance, trade finance and foreign exchange'), (87, 'Regulatory approvals and licences'),
      (88, 'Coordination between government agencies'), (89, 'Shortage of skilled labour')]
for i, (c, l) in enumerate(E1):
    add('e1_%d' % (i + 1), c, E, l, 'sev', order=SEV, grid='e1')
add('invest3', 90, F, 'Total investment FY2023-24 to FY2025-26, excluding land', ordinal=True)
F1 = [(92, 'Delayed an investment'), (93, 'Deferred an investment'), (94, 'Reduced the scale of a planned investment'),
      (95, 'Cancelled an investment'), (96, 'Used a less-preferred sourcing option'), (97, 'Lost export orders or competitiveness')]
add('f1', None, F, 'Effects of tariff or customs issues in the last 3 years', 'multi',
    cols=F1, none_label='None of these')
G1 = [(99, 'Cut tariffs on selected inputs'), (100, 'Simplify tariff bands, reduce escalation'),
      (101, 'Better HS guidance and advance rulings'), (102, 'Valuation transparency and dispute resolution'),
      (103, 'Faster clearance through digital reform'), (104, 'Simpler bonded, drawback and refund procedures'),
      (105, 'Advance notice and consultation on SRO changes'), (106, 'NBR, MoC, BTTC and IB coordination'),
      (107, 'Adjustment support for own sector')]
add('g1', None, G, 'Three most urgent reforms', 'multi', cols=G1, none_label='None selected')
add('confidence', 108, G, 'Confidence that tariff and customs rules stay the same over two years',
    order=['not at all confident', 'less confident', 'moderately confident', 'confident', 'very confident', 'DK'])
add('consult', 109, G, 'Effectiveness of government and business consultation on tariff changes',
    order=['very ineffective', 'ineffective', 'moderately effective', 'effective', 'very effective', 'no experience'])
add('overall', 110, G, 'Overall tariff and customs environment for investment',
    order=['Very constraining', 'Constraining', 'neutral', 'Supportive', 'Very supportive'])

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

for r in R:
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
    rec['grp'] = 'Domestic' if re.match(r'domestic|state', own, re.I) else 'Foreign or JV'
    rec['inzone'] = (rec['a']['zone'] or '').lower() != 'outside zones'
    rec['a']['age'] = (NOW.year - rec['a']['year']) if rec['a']['year'] else None
    rec['a']['workers'] = (rec['a']['perm'] or 0) + (rec['a']['temp'] or 0) if rec['a']['perm'] is not None else None
    # products
    imps = []
    for pc, vc, hc in [(29, 30, 31), (33, 34, 35), (37, 38, 39)]:
        if r[pc] in (None, ''): continue
        v, prob = money(r[vc])
        hs = re.sub(r'\D', '', str(r[hc] or ''))
        imps.append(dict(name=clean(r[pc]), value=v, raw=str(r[vc]).strip() if r[vc] is not None else '', hs=hs))
        if prob: issue(rec, 'Import value for "%s": %s (%s)' % (clean(r[pc]), prob.lower(), str(r[vc]).strip()), 'warn', 'Product value format')
        if hs and len(hs) < 4: issue(rec, 'HS code %s for "%s" is shorter than 4 digits' % (hs, clean(r[pc])), 'warn', 'HS code')
        if not hs: issue(rec, 'No HS code for imported "%s"' % clean(r[pc]), 'warn', 'HS code')
    exps = []
    for pc, hc, vc in [(52, 53, 54), (56, 57, 58)]:
        if r[pc] in (None, ''): continue
        v, prob = money(r[vc])
        hs = re.sub(r'\D', '', str(r[hc] or ''))
        exps.append(dict(name=clean(r[pc]), value=v, raw=str(r[vc]).strip() if r[vc] is not None else '', hs=hs))
        if prob: issue(rec, 'Export value for "%s": %s (%s)' % (clean(r[pc]), prob.lower(), str(r[vc]).strip()), 'warn', 'Product value format')
        if hs and len(hs) < 4: issue(rec, 'HS code %s for "%s" is shorter than 4 digits' % (hs, clean(r[pc])), 'warn', 'HS code')
    rec['imports'], rec['exports'] = imps, exps
    for i in range(3): rec['a']['imp_v%d' % (i + 1)] = imps[i]['value'] if i < len(imps) else None
    for i in range(2): rec['a']['exp_v%d' % (i + 1)] = exps[i]['value'] if i < len(exps) else None
    rec['suggest'] = clean(r[111]) or ''
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
    if rec['grp'] == 'Foreign or JV' and r[12] in (None, ''):
        issue(rec, 'Foreign ownership share is blank for a foreign or joint-venture firm', 'warn', 'Ownership share blank')
    if rec['a']['imp'] == 'no' and imps:
        issue(rec, 'Says no imports in 12 months but lists imported products', 'bad', 'Contradiction')
    if rec['a']['exp_dir'] == '0%' and rec['a']['exp_ind'] == '0%' and exps:
        issue(rec, 'Reports 0% exports but lists export products', 'warn', 'Contradiction')
    s = rec['suggest']
    if s.lower() in ('yes', 'no'):
        issue(rec, 'Answered "%s" but named no tariff line, SRO or procedure' % s, 'warn', 'Open answer')
    elif s and (len(s) < 5 or re.match(r'^(open record|hh+n+)', s, re.I)):
        issue(rec, 'Open answer looks like a placeholder: "%s"' % s, 'warn', 'Open answer')
    if rec['a']['perm'] is not None and rec['a']['perm'] == 0:
        issue(rec, 'Zero permanent workers', 'warn', 'Workers')
    key = re.sub(r'[^a-z]', '', rec['firm'].lower())
    if key in name_seen:
        issue(rec, 'Same firm name as record %s, possible duplicate' % name_seen[key], 'bad', 'Possible duplicate')
    else:
        name_seen[key] = rec['id']
    recs.append(rec)

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
NUMS = [('year', A, 'Year operations began in Bangladesh', 'year'), ('age', A, 'Years operating in Bangladesh', 'years'),
        ('perm', A, 'Permanent workers at present', 'workers'), ('temp', A, 'Temporary workers at present', 'workers'),
        ('workers', A, 'All workers (permanent + temporary)', 'workers'),
        ('imp_v1', B, 'Import value, top imported product 1', 'Tk'), ('imp_v2', B, 'Import value, top imported product 2', 'Tk'),
        ('imp_v3', B, 'Import value, top imported product 3', 'Tk'),
        ('exp_v1', B, 'Export value, main export product 1', 'Tk'), ('exp_v2', B, 'Export value, main export product 2', 'Tk')]
schema = [s for s in schema if s['type'] != 'num'] + [dict(key=k, section=s, label=l, type='num', unit=u) for k, s, l, u in NUMS]

versions = Counter(rc['version'] for rc in recs)
last_sub = max(str(r[133]) for r in R if r[133])
out = dict(title='Investor Survey 2026', updated=max(rc['start'] for rc in recs),
           last_submission_utc=last_sub, exported=(sys.argv[4] if len(sys.argv) > 4 else NOW.strftime('%Y-%m-%d %H:%M')),
           schema=schema, records=recs, issues=issues,
           versions=[dict(v=v, n=n, first=min(rc['date'] for rc in recs if rc['version'] == v),
                          last=max(rc['date'] for rc in recs if rc['version'] == v)) for v, n in versions.most_common()])
json.dump(out, open(OUT, 'w'), ensure_ascii=False, separators=(',', ':'))
print('records', len(recs), 'issues', len(issues))
for i in issues: print(' -', i['id'], i['issue'])
print(Counter(rc['grp'] for rc in recs))
