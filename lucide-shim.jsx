// Redesign: swaps Lucide icons for Google's Material Symbols (the prototype's icon set)
// without touching the app code. Every `import { X } from "lucide-react"` resolves here.
import { forwardRef } from "react";

const NAMES = {
  Clock: "schedule", CircleCheck: "check_circle", CircleX: "cancel", CircleAlert: "error",
  AlertTriangle: "warning", Paperclip: "attach_file", Image: "image", Circle: "radio_button_unchecked",
  Type: "title", Lock: "lock", BookOpen: "menu_book", Play: "play_arrow", RefreshCw: "autorenew",
  CircleHelp: "help", ExternalLink: "open_in_new", Check: "check", Copy: "content_copy",
  PanelLeftClose: "left_panel_close", Search: "search", X: "close", Filter: "center_focus_strong",
  Users: "group", Eye: "visibility", PanelLeftOpen: "left_panel_open", ChevronLeft: "chevron_left",
  ChevronRight: "chevron_right", Home: "home", RotateCcw: "restart_alt", Download: "download",
  FileText: "description", FileDown: "file_download", Plus: "add", Settings2: "settings",
  GripVertical: "drag_indicator", Pencil: "edit", Trash2: "delete", Upload: "upload",
  ArrowLeft: "arrow_back", Monitor: "desktop_windows", GitBranch: "fork_right", Bell: "notifications",
  AtSign: "alternate_email", List: "view_list", LayoutList: "view_agenda", Sun: "light_mode", Moon: "dark_mode", SlidersHorizontal: "tune", Link: "link", Ellipsis: "more_horiz", FileUp: "file_upload",
};

// Lucide icons are sized with Tailwind classes like "w-4 h-4" (16px). Material Symbols glyphs
// sit inside a padded box, so they're drawn ~15% larger to look the same size.
const sizeFrom = (cls = "") => {
  const m = String(cls).match(/(?:^|\s)w-(\d+(?:\.\d+)?)(?=\s|$)/) || String(cls).match(/(?:^|\s)w-\[(\d+)px\]/);
  if (!m) return 16;
  return m[0].includes("[") ? +m[1] : +m[1] * 4;
};

const make = (name, glyph) => {
  const C = forwardRef(({ className = "", strokeWidth, size, style, color, fill, ...rest }, ref) => {
    const px = size || sizeFrom(className);
    const wght = strokeWidth >= 2.5 ? 700 : 400;
    return (
      <span
        ref={ref}
        aria-hidden="true"
        className={`ms-ic ${className}`}
        style={{
          fontSize: Math.round(px * 1.18),
          width: px,
          height: px,
          fontVariationSettings: `'wght' ${wght}, 'FILL' 0, 'opsz' 20`,
          color,
          ...style,
        }}
        {...rest}
      >
        {glyph}
      </span>
    );
  });
  C.displayName = name;
  return C;
};

const exp = {};
for (const [k, v] of Object.entries(NAMES)) exp[k] = make(k, v);
export const {
  Clock, CircleCheck, CircleX, CircleAlert, AlertTriangle, Paperclip, Image, Circle, Type, Lock,
  BookOpen, Play, RefreshCw, CircleHelp, ExternalLink, Check, Copy, PanelLeftClose, Search, X,
  Filter, Users, Eye, PanelLeftOpen, ChevronLeft, ChevronRight, Home, RotateCcw, Download, FileText,
  FileDown, Plus, Settings2, GripVertical, Pencil, Trash2, Upload, ArrowLeft, Monitor, GitBranch,
  Bell, AtSign, List, LayoutList, Sun, Moon, SlidersHorizontal, Link, Ellipsis, FileUp,
} = exp;
