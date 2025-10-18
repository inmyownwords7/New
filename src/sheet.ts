


function extractCellValue(prop: any): string {
  if (!prop) return "";
  switch (prop.type) {
    case "title":
      return (prop.title || []).map((t: any) => t?.plain_text || "").join("");
    case "rich_text":
      return (prop.rich_text || []).map((t: any) => t?.plain_text || "").join("");
    case "email": return prop.email || "";
    case "phone_number": return prop.phone_number || "";
    case "url": return prop.url || "";
    case "date": return prop.date?.start || "";
    case "status": return prop.status?.name || "";
    case "select": return prop.select?.name || "";
    case "multi_select": return (prop.multi_select || []).map((o: any) => o?.name || "").join(", ");
    case "people": return (prop.people || []).map((p: any) => p?.name || p?.person?.email || "").join(", ");
    case "number": return (prop.number ?? "").toString();
    case "checkbox": return prop.checkbox ? "TRUE" : "FALSE";
    default: return "";
  }
}



