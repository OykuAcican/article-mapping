# -*- coding: utf-8 -*-
"""Link model elements to VMSW articles by hand and fill OVZ_Keynote.

Left: model elements grouped by category / family / type / system / size.
Right: the VMSW article list. Drag an article onto rows, apply, and every
choice is stored in a shared mapping file that pre-fills the next project.
"""
import os
import datetime

import clr
import System
from System import Tuple, DBNull, Action
from pyrevit import revit, DB, forms, script, HOST_APP
from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent

# The window stays open next to Revit, so the script engine must stay alive.
__persistentengine__ = True

for _asm in ('System.Data', 'System.Data.Common'):
    try:
        clr.AddReference(_asm)
    except Exception:
        pass
from System.Data import DataTable
from System.Collections.Generic import List
from System.Windows import DragDrop, DragDropEffects, DataFormats, SystemParameters
from System.Windows.Controls import (DataGridRow, TextBlock, CheckBox, StackPanel, Expander,
                                     GroupItem, DataGridCell, Control, ListBoxItem, TextBox,
                                     Button)
from System.Windows import Style, Setter, Visibility
from System.Windows.Media import SolidColorBrush, Color
from System.Windows.Data import PropertyGroupDescription, CollectionViewGroup
import System.Windows.Media
from System.Windows.Input import MouseButtonState, Key
from System.Windows.Media import VisualTreeHelper
from System.Windows import LogicalTreeHelper

from xlsx_reader import XlsxBook
from ovz_mapping import (make_key, load_articles, load_mapping, find_conflicts, DEFAULT_PARAM,
                         merge, write_mapping)

doc = revit.doc
uidoc = revit.uidoc
# Settings are kept under a fixed name. The first time, they are copied from
# the old "OVZ Keynote" button so nobody has to choose their files again.
config = script.get_config('OVZParamMapper')
_SETTINGS = ('ovz_article_list', 'ovz_mapping', 'ovz_grouping', 'ovz_target',
             'ovz_value_column', 'ovz_follow_selection', 'ovz_extra_columns')
if not config.get_option('ovz_article_list', ''):
    for _old in ('OVZKeynote', 'OVZ Keynote'):
        try:
            _old_cfg = script.get_config(_old)
            if _old_cfg.get_option('ovz_article_list', ''):
                for _k in _SETTINGS:
                    _v = _old_cfg.get_option(_k, None)
                    if _v is not None:
                        setattr(config, _k, _v)
                script.save_config()
                break
        except Exception:
            pass

HERE = os.path.dirname(__file__)
XAML_FILE = os.path.join(HERE, 'ui.xaml')
LOGO_FILE = os.path.normpath(os.path.join(HERE, '..', '..', '..', 'lib',
                                          'overzicht-octagon-blackred-rgb.png'))
P_INST = DEFAULT_PARAM            # default target parameter
TYPE_SUFFIX = ' (type)'
ART_CODE_COL = 'Art_nr.'
SHEET_NAME = 'alle delen'
FT_TO_MM = 304.8
ALL_CATS = 'All categories'
ALL_SUBS = 'All disciplines'
USER = os.environ.get('USERNAME', '')

BIP = DB.BuiltInParameter

OTHER = 'Other'
DISCIPLINES = ['Architecture', 'Structure', 'HVAC', 'Plumbing', 'Electrical',
               'Fire protection']
_DISC_CATS = {
    'Architecture': [
        'Walls', 'Floors', 'Ceilings', 'Roofs', 'Doors', 'Windows', 'Stairs',
        'Railings', 'Ramps', 'Columns', 'Curtain Panels', 'Curtain Wall Mullions',
        'Curtain Systems', 'Furniture', 'Furniture Systems', 'Casework',
        'Generic Models', 'Specialty Equipment', 'Rooms', 'Areas', 'Parking',
        'Planting', 'Site', 'Topography', 'Toposolid', 'Entourage', 'Mass',
        'Signage', 'Food Service Equipment', 'Medical Equipment', 'Hardscape',
        'Roads', 'Shaft Openings', 'Vertical Circulation', 'Temporary Structures'],
    'Structure': [
        'Structural Columns', 'Structural Framing', 'Structural Foundations',
        'Structural Rebar', 'Structural Connections', 'Structural Stiffeners',
        'Structural Trusses', 'Structural Beam Systems', 'Structural Area Reinforcement',
        'Structural Path Reinforcement', 'Structural Fabric Reinforcement',
        'Structural Tendons', 'Bearings', 'Abutments', 'Piers', 'Bridge Decks'],
    'HVAC': [
        'Ducts', 'Duct Fittings', 'Duct Accessories', 'Duct Insulations', 'Duct Linings',
        'Flex Ducts', 'Duct Placeholders', 'Air Terminals', 'Mechanical Equipment',
        'Mechanical Control Devices', 'MEP Ancillary Framing', 'Mechanical Equipment Sets'],
    'Plumbing': [
        'Pipes', 'Pipe Fittings', 'Pipe Accessories', 'Pipe Insulations', 'Flex Pipes',
        'Pipe Placeholders', 'Plumbing Fixtures', 'Plumbing Equipment'],
    'Electrical': [
        'Electrical Equipment', 'Electrical Fixtures', 'Lighting Fixtures',
        'Lighting Devices', 'Cable Trays', 'Cable Tray Fittings', 'Conduits',
        'Conduit Fittings', 'Communication Devices', 'Data Devices', 'Security Devices',
        'Telephone Devices', 'Nurse Call Devices', 'Audio Visual Devices',
        'Electrical Analytical Loads'],
    'Fire protection': ['Sprinklers', 'Fire Alarm Devices', 'Fire Protection'],
}
DISCIPLINE_OF = dict((c, d) for d, cats in _DISC_CATS.items() for c in cats)
# header texts use Revit's parameter names; g8-g10 are set per target in _build_model_table
MODEL_COLS = ['Category', 'Family', 'Type', 'System Type', 'System Abbreviation', 'Size',
              'Elements', 'Elements without value', 'OVZ_Keynote now', 'Article', 'Description',
              'System Classification', 'Shape']
# column order on screen (indexes into MODEL_COLS)
DISPLAY_ORDER = [0, 1, 2, 11, 3, 5, 6, 7, 8, 9, 10, 12]
# kept in the table (search, mapping key) but not shown: g4 = System Abbreviation
HIDDEN_MODEL_COLS = ('g4',)
COL_TIPS = {'g6': 'Number of elements in this row',
            'g7': 'Elements in this row whose parameter is empty',
            'g8': 'What is in Revit now for this parameter',
            'g9': 'Value you dragged from the Excel list. It is written to Revit when '
                  'you click Apply to model.',
            'g10': 'Article (row) in the Excel list the new value comes from'}
NONE_FIELD = '(none)'
# Category is always the first group level; these are the choices below it
# (the abbreviation is still read for the mapping key, it is just not offered here)
GROUP_FIELDS = ['Family', 'Type', 'System Classification', 'System Type',
                'Size', 'Diameter', 'Width', 'Height', 'Shape']
DEFAULT_GROUPING = ['Type', 'System Classification', 'Size']
# names used in saved settings before: old spelling, or the abbreviation (no longer offered)
OLD_FIELD_NAMES = {'System classification': 'System Classification',
                   'System type': 'System Type',
                   'System abbreviation': 'System Classification',
                   'System Abbreviation': 'System Classification'}
# between the fixed fields and the model's parameters in the Group by lists
# (disabled in ui.xaml, GroupChoiceItem)
SEPARATOR = unichr(0x2500) * 10
SHAPE_ORDER = {'Round': 0, 'Rectangular': 1, 'Oval': 2}
ART_COLS = ['Art_nr', 'Omschrijving', 'Unit', 'Discipline']
GROUP_KEY_COLS = ('k0', 'k1', 'k2', 'k3')
HIDDEN_COLS = ('id', 'state', 'cat', 'hasparam',
               'k0', 'k1', 'k2', 'k3', 's0', 's1', 's2', 's3')


# --------------------------------------------------------------------------
# Reading the model
# --------------------------------------------------------------------------
def element_name(el):
    for getter in (lambda e: e.Name,
                   lambda e: DB.Element.Name.__get__(e),
                   lambda e: DB.Element.Name.GetValue(e)):
        try:
            return getter(el) or ''
        except Exception:
            pass
    return ''


def param_text(el, bip):
    # Note: don't check HasValue - calculated parameters such as System
    # Abbreviation report HasValue = False even when Revit shows a value.
    p = el.get_Parameter(bip)
    if p is None:
        return ''
    try:
        if p.StorageType == DB.StorageType.String:
            return (p.AsString() or '').strip()
        return (p.AsValueString() or '').strip()
    except Exception:
        return ''


def _system_of(el):
    """The MEP system of a duct/pipe or of a family instance (via its connectors)."""
    try:
        sys = el.MEPSystem
        if sys is not None:
            return sys
    except Exception:
        pass
    try:
        for con in el.MEPModel.ConnectorManager.Connectors:
            if con.MEPSystem is not None:
                return con.MEPSystem
    except Exception:
        pass
    return None


def system_abbreviation(el):
    text = param_text(el, BIP.RBS_SYSTEM_ABBREVIATION_PARAM)
    if text:
        return text
    sys = _system_of(el)
    if sys is not None:
        try:
            sys_type = doc.GetElement(sys.GetTypeId())
            return (sys_type.Abbreviation or '').strip()
        except Exception:
            pass
    return ''


def param_mm(el, bip):
    p = el.get_Parameter(bip)
    if p is None or not p.HasValue:
        return 0
    return int(round(p.AsDouble() * FT_TO_MM))


def size_text(el):
    d = param_mm(el, BIP.RBS_PIPE_DIAMETER_PARAM)
    if d:
        return 'DN{}'.format(d)
    d = param_mm(el, BIP.RBS_CURVE_DIAMETER_PARAM)
    if d:
        return u'\u00d8{}'.format(d)
    w, h = param_mm(el, BIP.RBS_CURVE_WIDTH_PARAM), param_mm(el, BIP.RBS_CURVE_HEIGHT_PARAM)
    if w and h:
        return '{}x{}'.format(w, h)
    s = param_text(el, BIP.RBS_CALCULATED_SIZE)
    return s.replace(' mm', '').replace('mm', '').replace(' ', '')


def shape_and_dims(el, size):
    """Return (shape, diameter, width, height) in mm (0 = not applicable)."""
    d = param_mm(el, BIP.RBS_PIPE_DIAMETER_PARAM) or param_mm(el, BIP.RBS_CURVE_DIAMETER_PARAM)
    w, h = param_mm(el, BIP.RBS_CURVE_WIDTH_PARAM), param_mm(el, BIP.RBS_CURVE_HEIGHT_PARAM)
    shape = ''
    try:
        if isinstance(el, DB.Mechanical.Duct):
            shape = str(el.DuctType.Shape)          # Round / Rectangular / Oval
    except Exception:
        shape = ''
    if not shape:
        if d:
            shape = 'Round'
        elif w and h:
            shape = 'Rectangular'
        elif 'x' in size.lower():
            shape = 'Rectangular'
        elif u'\u00d8' in size or size.upper().startswith('DN'):
            shape = 'Round'
    if shape != 'Round':
        d = 0
    return shape, d, w if shape != 'Round' else 0, h if shape != 'Round' else 0


def lookup_text(el, el_type, name):
    for target in (el, el_type):
        if target is None:
            continue
        p = target.LookupParameter(name)
        if p is not None:
            if p.StorageType == DB.StorageType.String:
                return (p.AsString() or '').strip()
            return (p.AsValueString() or '').strip()
    return ''


class Group(object):
    def __init__(self, idx, key, cat, fam, typ, systype, abbr, size):
        self.idx = idx
        self.key = key
        self.category, self.family, self.type = cat, fam, typ
        self.systype, self.abbr, self.size = systype, abbr, size
        self.ids = []
        self.values = set()        # current values of the target parameter
        self.type_id = None
        self.source = ''           # article the new value came from
        self.empty = 0
        self.has_param = False
        self.extra = {}            # extra column -> set of values
        self.extra_values = {}     # Group by parameter -> value (rows are split by these)
        self.classification = ''
        self.shape = ''
        self.diameter = self.width = self.height = 0
        self.article = ''
        self.state = ''            # '', 'saved', 'manual', 'done'

    @property
    def system(self):
        return self.abbr or self.systype

    @property
    def current(self):
        vals = [v for v in self.values if v]
        if not vals:
            return ''
        return vals[0] if len(vals) == 1 and not self.empty else '<mixed>'


# Not physical model elements: views/sheets/cameras, materials, project info,
# links, imports, datums, lines, MEP system definitions (Duct/Piping Systems,
# electrical circuits)... (view-specific items like detail items are skipped
# separately)
SKIP_CLASSES = (DB.View, DB.Material, DB.ProjectInfo, DB.RevitLinkInstance,
                DB.ImportInstance, DB.Family, DB.Level, DB.Grid, DB.ReferencePlane,
                DB.SketchPlane, DB.CurveElement, DB.ProjectLocation, DB.BasePoint,
                DB.MEPSystem)


def split_target(target):
    if target.endswith(TYPE_SUFFIX):
        return target[:-len(TYPE_SUFFIX)], True
    return target, False


WRITABLE_STORAGE = (DB.StorageType.String, DB.StorageType.Integer, DB.StorageType.Double)


def _text_params(element, suffix, names):
    """Every editable parameter: text, number, integer and Yes/No.
    (Element references such as Material can't be filled from a text value.)
    names: name -> True when it is a shared parameter."""
    try:
        for p in element.Parameters:
            if p.StorageType in WRITABLE_STORAGE and not p.IsReadOnly:
                name = p.Definition.Name + suffix
                names[name] = names.get(name, False) or bool(p.IsShared)
    except Exception:
        pass


def _all_params(element, suffix, names):
    """Every parameter name, any kind and read-only too (for Group by)."""
    try:
        for p in element.Parameters:
            names.add(p.Definition.Name + suffix)
    except Exception:
        pass


def split_value(el, el_type, name):
    """Value of a Group by parameter ('<name> (type)' = from the element type)."""
    pname, on_type = split_target(name)
    target = el_type if on_type else el
    p = target.LookupParameter(pname) if target is not None else None
    return read_value(p) if p is not None else ''


def read_value(p):
    """Value of any parameter as text, the way Revit shows it."""
    try:
        if p.StorageType == DB.StorageType.String:
            return (p.AsString() or '').strip()
        if not p.HasValue:
            return ''
        return (p.AsValueString() or '').strip()
    except Exception:
        return ''


YES_WORDS = ('yes', 'y', 'true', '1', 'ja', 'oui', 'evet', 'x')
NO_WORDS = ('no', 'n', 'false', '0', 'nee', 'non', 'hayir', u'hay\u0131r')


def _same_number(a, b):
    """'300 mm' and '300' or '2,5' and '2.5' count as the same value."""
    import re
    def num(text):
        m = re.search(r'-?\d+(?:[.,]\d+)?', text or '')
        return float(m.group(0).replace(',', '.')) if m else None
    x, y = num(a), num(b)
    return x is not None and y is not None and abs(x - y) < 1e-6


def write_value(p, raw):
    """Write text into any parameter. Returns True if the value changed.
    Raises ValueError when the text doesn't fit the parameter."""
    st = p.StorageType
    if st == DB.StorageType.String:
        if (p.AsString() or '') == raw:
            return False
        p.Set(raw)
        return True
    if st == DB.StorageType.Integer:
        low = raw.strip().lower()
        if low in YES_WORDS:
            v = 1
        elif low in NO_WORDS:
            v = 0
        else:
            try:
                v = int(float(raw.replace(',', '.')))
            except ValueError:
                raise ValueError(u'"{}" is not a whole number or yes/no'.format(raw))
        if p.AsInteger() == v:
            return False
        p.Set(v)
        return True
    if st == DB.StorageType.Double:
        before = p.AsDouble()
        # read in project units ("300" = 300 mm in a mm project, "300 mm" works too)
        ok = p.SetValueString(raw)
        if not ok and ('.' in raw or ',' in raw):
            ok = p.SetValueString(raw.replace('.', '#').replace(',', '.').replace('#', ','))
        if not ok:
            raise ValueError(u'"{}" is not a number'.format(raw))
        return abs(p.AsDouble() - before) > 1e-9
    raise ValueError('this kind of parameter cannot be filled from text')


def gather_groups(extra_names, target=P_INST, split_names=()):
    """Group the model elements. Also returns every editable text parameter
    found (instance names, and type names with ' (type)') -> is shared, and every parameter
    name of any kind (for Group by). Rows are also split by the values of
    split_names; the mapping key (g.key) does not include them."""
    pname, on_type = split_target(target)
    param_names = {}               # name -> is shared
    all_names = set()
    seen_types = set()
    groups, by_key = [], {}
    els = DB.FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
    for el in els:
        cat = el.Category
        if cat is None or cat.CategoryType != DB.CategoryType.Model:
            continue
        if isinstance(el, SKIP_CLASSES) or el.ViewSpecific:
            continue
        tid = el.GetTypeId()
        el_type = doc.GetElement(tid) if tid and tid != DB.ElementId.InvalidElementId else None
        if on_type:
            p = el_type.LookupParameter(pname) if el_type is not None else None
        else:
            p = el.LookupParameter(pname)
        if el_type is None and p is None:
            continue
        fam = ''
        if el_type is not None:
            try:
                fam = el_type.FamilyName or ''
            except Exception:
                fam = ''
        typ = element_name(el_type) if el_type is not None else ''
        systype = (param_text(el, BIP.RBS_PIPING_SYSTEM_TYPE_PARAM)
                   or param_text(el, BIP.RBS_DUCT_SYSTEM_TYPE_PARAM))
        abbr = system_abbreviation(el)
        size = size_text(el)
        key = make_key(cat.Name, fam, typ, abbr or systype, size)
        split_vals = tuple(split_value(el, el_type, n) for n in split_names)
        row_key = (key,) + split_vals
        g = by_key.get(row_key)
        if g is None:
            g = Group(len(groups), key, cat.Name, fam, typ, systype, abbr, size)
            g.extra_values = dict(zip(split_names, split_vals))
            g.classification = param_text(el, BIP.RBS_SYSTEM_CLASSIFICATION_PARAM)
            g.shape, g.diameter, g.width, g.height = shape_and_dims(el, size)
            g.type_id = el_type.Id if el_type is not None else None
            _text_params(el, '', param_names)
            _all_params(el, '', all_names)
            if el_type is not None:
                _text_params(el_type, TYPE_SUFFIX, param_names)
                if str(tid) not in seen_types:
                    seen_types.add(str(tid))
                    _all_params(el_type, TYPE_SUFFIX, all_names)
            groups.append(g)
            by_key[row_key] = g
        g.ids.append(el.Id)
        if p is not None:
            g.has_param = True
            value = read_value(p)
            g.values.add(value)
            if not value:
                g.empty += 1
        for name in extra_names:
            g.extra.setdefault(name, set()).add(lookup_text(el, el_type, name))
    return groups, param_names, all_names


def _num_key(text):
    """'Ø1000' > 'Ø125': sort sizes by their numbers, not as text."""
    import re
    nums = re.findall(r'\d+', text or '')
    return ' '.join('%08d' % int(n) for n in nums) + ' ' + (text or '')


def group_value(g, field):
    """(label, sort key) of a row for one grouping field."""
    if field == 'Category':
        v = g.category
    elif field == 'Family':
        v = g.family
    elif field == 'Type':
        v = g.type
    elif field == 'System Classification':
        v = g.classification
    elif field == 'System Type':
        v = g.systype
    elif field == 'System Abbreviation':
        v = g.abbr
    elif field == 'Shape':
        v = g.shape
        return (v or '(no shape)', '%d %s' % (SHAPE_ORDER.get(v, 9), v))
    elif field == 'Size':
        v = g.size
        return (u'Size: {}'.format(v) if v else '(no size)',
                '%d %s' % (SHAPE_ORDER.get(g.shape, 9), _num_key(v)))
    elif field in ('Diameter', 'Width', 'Height'):
        n = {'Diameter': g.diameter, 'Width': g.width, 'Height': g.height}[field]
        if not n:
            return ('(no {})'.format(field.lower()), '99999999')
        return ((u'\u00d8{}' if field == 'Diameter' else '{}').format(n), '%08d' % n)
    else:
        # any other parameter chosen in Group by: numbers first (by value), then text
        v = g.extra_values.get(field, '')
        if not v:
            return ('(empty)', '2')
        if v[0].isdigit():
            return (v, '0 ' + _num_key(v))
        return (v, '1 ' + v.lower())
    return (v or '(empty)', (v or '~').lower())


# --------------------------------------------------------------------------
# Revit may only be changed when it is ready ("API context"). A window that
# stays open asks Revit to run an action for it through an ExternalEvent.
# --------------------------------------------------------------------------
class RevitAction(IExternalEventHandler):
    def __init__(self):
        self.func = None

    def Execute(self, uiapp):
        func, self.func = self.func, None
        if func is not None:
            func(uiapp)

    def GetName(self):
        return 'OVZ Param Mapper'


ACTION = RevitAction()
EVENT = ExternalEvent.Create(ACTION)


# --------------------------------------------------------------------------
# Window
# --------------------------------------------------------------------------
def find_ancestor(obj, cls):
    while obj is not None:
        if isinstance(obj, cls):
            return obj
        parent = None
        try:
            parent = VisualTreeHelper.GetParent(obj)
        except Exception:
            pass
        if parent is None:
            try:
                parent = LogicalTreeHelper.GetParent(obj)
            except Exception:
                parent = None
        obj = parent
    return None


def esc(text):
    """Make text safe inside a DataView RowFilter LIKE pattern."""
    for ch in "[]*%":
        text = text.replace(ch, '')
    return text.replace("'", "''")


class AutoCompletePicker(object):
    """A TextBox with a ▾ button and a list popup: type part of a name or open
    the whole list. The popup has StaysOpen True; this class closes it itself
    (choice, Escape, focus leaves, window moved or deactivated) and then puts
    back the current name, so no half-typed text is left behind."""

    def __init__(self, window, box, arrow, popup, listbox, on_choose,
                 empty_text='No matching item'):
        self.window, self.box, self.arrow = window, box, arrow
        self.popup, self.list = popup, listbox
        self.on_choose = on_choose
        self.empty_text = empty_text
        self.names = []            # priority items first, otherwise in the given order
        self.current = ''
        self._setting = False      # box text set by code, not typed
        # all pickers of this window, so opening one closes the others
        if not hasattr(window, '_pickers'):
            window._pickers = []
        window._pickers.append(self)
        for ev, handler in ((box.PreviewMouseLeftButtonDown, self._on_mouse_down),
                            (box.PreviewMouseLeftButtonUp, self._on_click),
                            (box.GotKeyboardFocus, self._on_got_focus),
                            (box.TextChanged, self._on_text_changed),
                            (box.PreviewKeyDown, self._on_key),
                            (box.LostKeyboardFocus, self._on_lost_focus),
                            (arrow.Click, self._on_arrow),
                            (listbox.PreviewMouseLeftButtonUp, self._on_list_click),
                            (window.PreviewMouseDown, self._on_window_mouse_down),
                            (window.LocationChanged, self._on_window_moved),
                            # (a click outside the tool window, e.g. in the Revit model)
                            (window.Deactivated, self._on_window_moved)):
            ev += handler

    # -- called by the window ---------------------------------------------------------
    def set_items(self, items, current):
        """items: list of (name, is_priority). Priority items come first."""
        ordered = sorted(enumerate(items), key=lambda x: (not x[1][1], x[0]))
        self.names = [name for _, (name, _p) in ordered]
        self.set_current(current)

    def set_current(self, name):
        self.current = name
        self._set_text(name)

    # -- internals ----------------------------------------------------------------------
    def _set_text(self, text):
        self._setting = True
        self.box.Text = text
        self._setting = False

    def _later(self, func):
        """Run func after the current mouse/focus event has finished."""
        self.window.Dispatcher.BeginInvoke(Action(func))

    def _matches(self, text):
        """Names containing text: exact match, then starts with, then contains."""
        t = (text or '').strip().lower()
        if not t:
            return list(self.names)
        exact, start, inside = [], [], []
        for n in self.names:                    # already priority first
            low = n.lower()
            if low == t:
                exact.append(n)
            elif low.startswith(t):
                start.append(n)
            elif t in low:
                inside.append(n)
        return exact + start + inside

    def _show(self, names, select=None):
        lst = self.list
        if names:
            lst.ItemsSource = names
            lst.SelectedIndex = names.index(select) if select in names else 0
        else:
            item = ListBoxItem()
            item.Content = self.empty_text
            item.IsEnabled = False
            lst.ItemsSource = [item]
            lst.SelectedIndex = -1
        for other in self.window._pickers:
            if other is not self and other.popup.IsOpen:
                other._close()
        self.popup.IsOpen = True
        if lst.SelectedItem is not None:
            lst.ScrollIntoView(lst.SelectedItem)

    def _show_full(self):
        self._show(list(self.names), self.current)

    def _open_full(self):
        if not self.popup.IsOpen:
            self._show_full()

    def _close(self):
        """Close the list and put back the current name."""
        self.popup.IsOpen = False
        self._set_text(self.current)

    def _move(self, step):
        if not self.popup.IsOpen:
            text = self.box.Text
            if text == self.current:
                self._show_full()
            else:
                self._show(self._matches(text))
            return
        lst = self.list
        if not lst.Items.Count or not isinstance(lst.Items[0], str):
            return                              # "No matching ..."
        i = max(0, min(lst.Items.Count - 1, lst.SelectedIndex + step))
        lst.SelectedIndex = i
        lst.ScrollIntoView(lst.Items[i])

    def _choose_highlighted(self):
        item = self.list.SelectedItem
        if self.popup.IsOpen and isinstance(item, str):
            self._choose(item)
            return True
        return False

    def _choose(self, name):
        self.popup.IsOpen = False
        self._set_text(name)
        self.box.SelectAll()
        if name == self.current:
            return
        self.current = name
        self.on_choose(name)

    # -- events ---------------------------------------------------------------------------
    def _on_mouse_down(self, sender, e):
        # first click: take the focus here, so GotFocus can select all text
        # (otherwise the click itself would put the caret and undo the selection)
        if not self.box.IsKeyboardFocusWithin:
            self.box.Focus()
            e.Handled = True

    def _on_click(self, sender, e):
        # (when the list is already open, a click only moves the caret)
        self._later(self._open_full)

    def _on_got_focus(self, sender, e):
        self.box.SelectAll()
        self._later(self._open_full)

    def _on_text_changed(self, sender, e):
        if self._setting or not self.box.IsKeyboardFocusWithin:
            return
        self._show(self._matches(self.box.Text))

    def _on_key(self, sender, e):
        if e.Key == Key.Down:
            self._move(1)
            e.Handled = True
        elif e.Key == Key.Up:
            self._move(-1)
            e.Handled = True
        elif e.Key == Key.Enter:
            self._choose_highlighted()
            e.Handled = True
        elif e.Key == Key.Tab:
            self._choose_highlighted()          # Tab still moves on to the next field
        elif e.Key == Key.Escape:
            self._close()
            self.box.SelectAll()
            e.Handled = True

    def _on_list_click(self, sender, e):
        item = find_ancestor(e.OriginalSource, ListBoxItem)
        if item is not None and item.IsEnabled and isinstance(item.Content, str):
            self._choose(item.Content)
            e.Handled = True

    def _on_arrow(self, sender, e):
        """Toggle the full list."""
        if self.popup.IsOpen:
            self.popup.IsOpen = False
            return
        self.box.Focus()
        self.box.SelectAll()
        self._later(self._open_full)

    def _on_lost_focus(self, sender, e):
        # check once the new focus is known; never leave half-typed text behind
        def check():
            if not (self.box.IsKeyboardFocusWithin or self.list.IsKeyboardFocusWithin):
                self._close()
        self._later(check)

    def _on_window_moved(self, sender, e):
        # the list would stay behind on screen (StaysOpen), so close it
        if self.popup.IsOpen:
            self._close()

    def _on_window_mouse_down(self, sender, e):
        """A click anywhere else in the window closes the list. (Clicks in the
        list itself never get here: the popup is its own window.) Not handled,
        so the click still does what it normally does."""
        if not self.popup.IsOpen:
            return
        src = e.OriginalSource
        if find_ancestor(src, TextBox) is self.box or find_ancestor(src, Button) is self.arrow:
            return                  # the box handles its own clicks, the arrow toggles
        self._close()


class OvzWindow(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self, XAML_FILE)
        if os.path.exists(LOGO_FILE):
            from System.Windows.Media.Imaging import BitmapImage
            self.logo_img.Source = BitmapImage(System.Uri(LOGO_FILE))
        self.groups = []
        self.model_table = None
        self.art_table = None
        self.desc = {}
        self.mapping = {}          # what is in the shared file (as last read)
        self.snapshot = {}         # the file as it was when we read it
        self.changes = {}          # key -> record or None (removed)
        # what was given to each element in this session: (target, element id) ->
        # (value, article, order); '' = removed. Rows keep their own value even
        # when several rows share one mapping key, also after the model is re-read.
        self._assigned = {}
        self._seq = 0
        self._split = []           # Group by parameters the rows were split by
        self.map_path = ''
        self.extra_names = []
        self._drag_start = None
        self._headers = {}
        self._rows = {}
        self._closing_ok = False
        self._busy = False
        self.follow_cb.IsChecked = config.get_option('ovz_follow_selection', True)
        self.follow_cb.Click += self.on_follow_toggle
        self.target = config.get_option('ovz_target', P_INST) or P_INST
        self.value_col = config.get_option('ovz_value_column', ART_CODE_COL) or ART_CODE_COL
        self.art_list = []
        self._art_col = {}
        self.art_columns = []
        self.art_values = {}
        # autocomplete pickers: Fill parameter (left) and Drag value from column (right)
        self.target_picker = AutoCompletePicker(
            self, self.target_box, self.target_arrow, self.target_popup, self.target_list,
            self._on_target_chosen, 'No matching parameter')
        self.value_picker = AutoCompletePicker(
            self, self.value_box, self.value_arrow, self.value_popup, self.value_list,
            self._on_value_col_chosen, 'No matching column')
        self._el_to_group = {}
        self._own_selection = set()
        self._sel_callback = None
        self._loading_grouping = True
        self._group_cbs = [self.group1_cb, self.group2_cb, self.group3_cb]
        saved = (config.get_option('ovz_grouping', '') or '').split(';')
        # older settings had 4 levels starting with Category: use the default then
        if len(saved) != 3 or saved[0] == 'Category':
            saved = DEFAULT_GROUPING
        saved = [OLD_FIELD_NAMES.get(v, v) for v in saved]
        # until the model is read, offer the saved parameters so they can be selected
        choices = self._group_choices(v for v in saved if v)
        for cb, value in zip(self._group_cbs, saved):
            cb.ItemsSource = choices
            cb.SelectedItem = value if value in choices else NONE_FIELD
            cb.SelectionChanged += self.on_grouping_changed
        self._loading_grouping = False
        self._cat_boxes = []       # category CheckBoxes in the popup
        self._disc_boxes = []      # discipline CheckBoxes (Tag = its categories)
        self._disc_name = []

        self.doc_name.Text = doc.Title

        for ev, handler in ((self.list_btn.Click, self.on_list_browse),
                            (self.map_open_btn.Click, self.on_map_open),
                            (self.map_new_btn.Click, self.on_map_new),
                            (self.select_btn.Click, self.on_select),
                            (self.isolate_btn.Click, self.on_isolate),
                            (self.clear_btn.Click, self.on_clear),
                            (self.assign_btn.Click, self.on_assign_button),
                            (self.unfilled_btn.Click, self.on_unfilled),
                            (self.colour_btn.Click, self.on_colour),
                            (self.uncolour_btn.Click, self.on_uncolour),
                            (self.save_btn.Click, self.on_save),
                            (self.saveas_btn.Click, self.on_save_as),
                            (self.apply_btn.Click, self.on_apply),
                            (self.close_btn.Click, self.on_close_click),
                            (self.refresh_btn.Click, self.on_refresh),
                            (self.cat_all_btn.Click, self.on_cat_all),
                            (self.cat_none_btn.Click, self.on_cat_none),
                            (self.collapse_all_btn.Click, self.on_collapse_all),
                            (self.expand_all_btn.Click, self.on_expand_all),
                            (self.model_search.TextChanged, self.on_model_filter),
                            (self.only_empty_cb.Click, self.on_model_filter),
                            (self.only_unassigned_cb.Click, self.on_model_filter),
                            (self.hide_noparam_cb.Click, self.on_hide_noparam),
                            (self.art_search.TextChanged, self.on_art_filter),
                            (self.sub_cb.SelectionChanged, self.on_art_filter),
                            (self.model_grid.AutoGeneratingColumn, self.on_auto_column),
                            (self.art_grid.AutoGeneratingColumn, self.on_auto_column),
                            (self.model_grid.DragOver, self.on_drag_over),
                            (self.model_grid.Drop, self.on_drop),
                            (self.model_grid.AutoGeneratedColumns, self.on_model_columns_ready),
                            (self.model_grid.PreviewMouseRightButtonUp, self.on_group_right_click),
                            (self.art_grid.PreviewMouseLeftButtonDown, self.on_art_mouse_down),
                            (self.art_grid.MouseMove, self.on_art_mouse_move),
                            (self.art_grid.MouseDoubleClick, self.on_art_double_click),
                            (self.Closing, self.on_closing),
                            (self.Closed, lambda s, e: self._unsubscribe_selection())):
            ev += handler

        path = config.get_option('ovz_article_list', '')
        if path and os.path.exists(path):
            self.load_articles(path)
        path = config.get_option('ovz_mapping', '')
        if path and os.path.exists(path):
            self.load_mapping_file(path)
        self.load_model()
        self._subscribe_selection()

    # -- loading ---------------------------------------------------------------
    def load_articles(self, path):
        try:
            book = XlsxBook(path)
            sheet = SHEET_NAME if SHEET_NAME in book.sheet_names else book.sheet_names[0]
            (self.art_list, self.desc, dups,
             self.art_columns, self.art_values) = load_articles(book.rows(sheet))
        except Exception as ex:
            forms.alert('Could not read the article list. If it is open in Excel, '
                        'save it and try again.\n\n{}'.format(ex))
            return
        if self.value_col not in self.art_columns:
            self.value_col = self.art_columns[0] if self.art_columns else ART_CODE_COL
        # Excel order, no priority items
        self.value_picker.set_items([(c, False) for c in self.art_columns], self.value_col)
        subs = sorted(set(a[3] for a in self.art_list if a[3]))
        self.sub_cb.ItemsSource = [ALL_SUBS] + subs
        self.sub_cb.SelectedIndex = 0
        self._build_art_table()
        self.list_box.Text = path
        config.ovz_article_list = path
        script.save_config()
        self._refresh_descriptions()

    def _build_art_table(self):
        """The whole article list, every column. The chosen value column is marked."""
        dt = DataTable()
        self._art_col = {}                      # Excel column name -> table column
        for i, name in enumerate(self.art_columns):
            cname = 'x{}'.format(i)
            dt.Columns.Add(cname)
            self._headers[cname] = name
            self._art_col[name] = cname
        for code, desc, unit, sub in self.art_list:
            vals = self.art_values.get(code, {})
            dr = dt.NewRow()
            for name, cname in self._art_col.items():
                dr[cname] = vals.get(name, '')
            dt.Rows.Add(dr)
        self.art_table = dt
        self.art_grid.ItemsSource = dt.DefaultView
        self._art_filter()

    def _on_value_col_chosen(self, name):
        self.value_col = name
        config.ovz_value_column = self.value_col
        script.save_config()
        self._build_art_table()

    def load_mapping_file(self, path):
        try:
            self.mapping = load_mapping(path)
        except Exception as ex:
            forms.alert('Could not read the mapping file.\n\n{}'.format(ex))
            return
        self.snapshot = dict(self.mapping)
        self.changes = {}
        self._assigned = {}
        self.map_path = path
        self.map_box.Text = '{}  ({} rules)'.format(path, len(self.mapping))
        config.ovz_mapping = path
        script.save_config()
        self._apply_mapping_to_groups(reset=True)

    def load_model(self):
        split = self._split_names()
        self.groups, names, all_names = gather_groups(self.extra_names, self.target, split)
        self._el_to_group = dict((str(eid), g.idx) for g in self.groups for eid in g.ids)
        self._fill_target_list(names)
        self._fill_group_lists(all_names)
        self._split = self._split_names()
        self._apply_mapping_to_groups(reset=True, refresh=False)
        self._build_model_table()
        self._build_cat_list()
        self._set_status()

    def _apply_mapping_to_groups(self, reset=False, refresh=True):
        for g in self.groups:
            key = self._mkey(g)
            given = self._row_assignment(g)
            if given is not None:
                # given (or removed) in this session: the row keeps its own value,
                # even if another row with the same mapping key got a different one
                g.article, g.source = given[0], given[1]
                if not g.article:
                    g.state = ''
                elif key in self.changes:
                    g.state = 'manual'
                else:
                    g.state = 'done' if g.current == g.article else 'saved'
                continue
            if key in self.changes and self.changes[key] is None:
                g.article, g.source, g.state = '', '', ''
                continue
            # (a rule changed through another row with this key is not copied
            # here: this row keeps what the mapping file has)
            rec = self.mapping.get(key)
            if not rec and g.abbr and g.systype:
                # rule saved with the system type name (before the abbreviation
                # was read): use it, and re-save it under the abbreviation
                old = self.mapping.get(self._mkey(g, g.systype))
                if old:
                    rec = dict(old)
                    rec['System'] = g.abbr
                    if key not in self.changes:
                        self.changes[key] = rec
                    g.article, g.source, g.state = rec['Value'], rec['Art_nr'], 'saved'
                    continue
            if rec:
                g.article, g.source = rec['Value'], rec['Art_nr']
                g.state = 'done' if g.current == g.article else 'saved'
            elif reset:
                g.article, g.source, g.state = '', '', ''
        if refresh and self.model_table is not None:
            for g in self.groups:
                self._update_row(g)
            self._set_status()

    # -- tables ------------------------------------------------------------------
    def _build_model_table(self):
        dt = DataTable()
        for name in HIDDEN_COLS:
            if name in GROUP_KEY_COLS:      # (label, count text) tuples for the headers
                dt.Columns.Add(name, clr.GetClrType(System.Object))
            else:
                dt.Columns.Add(name)
        cols = list(MODEL_COLS) + self.extra_names
        cols[8] = u'{} (current value in model)'.format(self.target)
        cols[9] = u'{} (new value, dragged from Excel)'.format(self.target)
        cols[10] = 'Source article (Excel)'
        for i, h in enumerate(cols):
            dt.Columns.Add('g{}'.format(i))
            self._headers['g{}'.format(i)] = h
        self.model_table = dt
        self._rows = {}
        for g in self.groups:
            dr = dt.NewRow()
            self._rows[g.idx] = dr
            dr['id'] = str(g.idx)
            dr['cat'] = g.category
            dr['hasparam'] = '1' if g.has_param else '0'
            dt.Rows.Add(dr)
            self._fill_row(dr, g)
        self._fill_group_columns()
        self.model_grid.ItemsSource = dt.DefaultView
        self._apply_grouping()
        self._model_filter()

    def _fill_row(self, dr, g):
        vals = [g.category, g.family, g.type, g.systype, g.abbr, g.size,
                len(g.ids), g.empty if g.has_param else '-',
                g.current if g.has_param else 'no parameter',
                g.article, self._source_text(g), g.classification, g.shape]
        for name in self.extra_names:
            s = [v for v in g.extra.get(name, set())]
            vals.append(s[0] if len(s) == 1 else ('<varies>' if s else ''))
        for i, v in enumerate(vals):
            dr['g{}'.format(i)] = str(v)
        dr['state'] = g.state

    def _source_text(self, g):
        code = g.source or (g.article if g.article in self.desc else '')
        if not code:
            return ''
        return u'{}  {}'.format(code, self.desc.get(code, ''))

    def _update_row(self, g):
        dr = self._rows.get(g.idx)
        if dr is not None:
            self._fill_row(dr, g)

    def _refresh_descriptions(self):
        if self.model_table is not None:
            for g in self.groups:
                self._update_row(g)

    # -- target parameter --------------------------------------------------------------
    def _fill_target_list(self, names):
        """names: parameter name -> True for a shared parameter."""
        shared = dict(names)
        shared.setdefault(P_INST, False)
        shared.setdefault(self.target, False)
        # shared parameters first (priority), then the others, each A-Z
        items = [(n, shared[n]) for n in sorted(shared, key=lambda n: n.lower())]
        self.target_picker.set_items(items, self.target)
        self.target_info.Text = '{} parameters'.format(len(items))

    def _on_target_chosen(self, name):
        self.target = name
        config.ovz_target = name
        script.save_config()

        def run():
            self.load_model()
            self._set_status('Now filling "{}"'.format(self.target))
        self._run(run)

    def _mkey(self, g, system=None):
        return make_key(g.category, g.family, g.type,
                        g.system if system is None else system, g.size, self.target)

    def _record(self, g, value, source):
        return {'Category': g.category, 'Family': g.family, 'Type': g.type,
                'System': g.system, 'Size': g.size, 'Parameter': self.target,
                'Value': value, 'Art_nr': source,
                'Omschrijving': self.desc.get(source, ''), 'Updated by': USER,
                'Updated on': datetime.date.today().isoformat(), 'Project': doc.Title}

    # -- grouping (like the Sorting/Grouping tab of a schedule) ---------------------
    def _grouping(self):
        out = ['Category']
        for cb in self._group_cbs:
            f = cb.SelectedItem
            if f and f not in (NONE_FIELD, SEPARATOR) and f not in out:
                out.append(f)
        return out

    def _split_names(self):
        """The chosen Group by parameters that are not fixed fields."""
        return [f for f in self._grouping()[1:] if f not in GROUP_FIELDS]

    def _group_choices(self, param_names):
        """Fixed fields, a separator, then the model's parameters A-Z."""
        skip = set(GROUP_FIELDS + ['Category', NONE_FIELD, SEPARATOR])
        params = sorted(set(n for n in param_names if n not in skip), key=lambda n: n.lower())
        return [NONE_FIELD] + GROUP_FIELDS + [SEPARATOR] + params

    def _fill_group_lists(self, param_names):
        """Offer every parameter found on the elements; keep the choice if it still exists."""
        choices = self._group_choices(param_names)
        self._loading_grouping = True
        for cb in self._group_cbs:
            current = cb.SelectedItem
            cb.ItemsSource = choices
            cb.SelectedItem = current if current in choices else NONE_FIELD
        self._loading_grouping = False

    def _fill_group_columns(self):
        fields = self._grouping()
        # element counts per group prefix, for the header text
        counts = {}
        labels = {}
        for g in self.groups:
            vals = [group_value(g, f) for f in fields]
            labels[g.idx] = vals
            for lvl in range(len(fields)):
                pre = tuple(v[0] for v in vals[:lvl + 1])
                counts[pre] = counts.get(pre, 0) + len(g.ids)
        for g in self.groups:
            dr = self._rows[g.idx]
            vals = labels[g.idx]
            for lvl in range(4):
                if lvl < len(fields):
                    pre = tuple(v[0] for v in vals[:lvl + 1])
                    # the header shows the label, then (category only) the count
                    # in grey; equal tuples (compared by value) fall in the same group
                    label = vals[0][0].upper() if lvl == 0 else vals[lvl][0]
                    n = counts[pre]
                    if lvl > 0:
                        count_text = ''         # sub-groups: the Elements column has the count
                    else:
                        count_text = '(1 element)' if n == 1 else '({} elements)'.format(n)
                    dr['k{}'.format(lvl)] = Tuple.Create(label, count_text)
                    dr['s{}'.format(lvl)] = vals[lvl][1]
                else:
                    dr['k{}'.format(lvl)] = DBNull.Value
                    dr['s{}'.format(lvl)] = ''

    def _apply_grouping(self):
        fields = self._grouping()
        # sort on the table itself (no extra .NET types needed, works on .NET 8)
        sort_cols = ['s{}'.format(lvl) for lvl in range(len(fields))] + ['g0', 'g1', 'g2', 'g5']
        self.model_table.DefaultView.Sort = ', '.join(c + ' ASC' for c in sort_cols)
        items = self.model_grid.Items
        items.GroupDescriptions.Clear()
        for lvl in range(len(fields)):
            items.GroupDescriptions.Add(PropertyGroupDescription('k{}'.format(lvl)))

    def on_grouping_changed(self, sender, e):
        if self.model_table is None or self._loading_grouping:
            return
        config.ovz_grouping = ';'.join(cb.SelectedItem or NONE_FIELD for cb in self._group_cbs)
        script.save_config()
        if sorted(self._split_names()) != sorted(self._split):
            # rows are split by these parameters' values: read the model again
            def run():
                self.load_model()
                self._set_status('Model read again for the new grouping')
            self._run(run)
            return
        self._fill_group_columns()
        self._apply_grouping()
        self._model_filter()

    def _category_expanders(self):
        """The Expander of every category header that WPF has built (on screen).
        Only categories have one; sub-groups are plain header + rows."""
        out = []
        todo = [self.model_grid]
        while todo:
            obj = todo.pop()
            for i in range(VisualTreeHelper.GetChildrenCount(obj)):
                child = VisualTreeHelper.GetChild(obj, i)
                if not isinstance(child, GroupItem):
                    todo.append(child)
                    continue
                # a category: don't look inside, its sub-groups have no Expander
                if (VisualTreeHelper.GetChildrenCount(child)
                        and find_ancestor(VisualTreeHelper.GetParent(child), GroupItem) is None):
                    exp = VisualTreeHelper.GetChild(child, 0)
                    if isinstance(exp, Expander):
                        out.append(exp)
        return out

    def on_collapse_all(self, sender, e):
        for exp in self._category_expanders():
            exp.IsExpanded = False

    def on_expand_all(self, sender, e):
        for exp in self._category_expanders():
            exp.IsExpanded = True

    def _groups_under(self, view_group):
        out = []
        for item in view_group.Items:
            if isinstance(item, CollectionViewGroup):
                out.extend(self._groups_under(item))
            else:
                out.append(self.groups[int(item.Row['id'])])
        return out

    def _rows_under(self, view_group):
        out = []
        for item in view_group.Items:
            if isinstance(item, CollectionViewGroup):
                out.extend(self._rows_under(item))
            else:
                out.append(item)
        return out

    def on_group_right_click(self, sender, e):
        """Right-click a group header: select every row in that group."""
        if find_ancestor(e.OriginalSource, DataGridRow) is not None:
            return
        gi = find_ancestor(e.OriginalSource, GroupItem)
        if gi is None or not isinstance(gi.DataContext, CollectionViewGroup):
            return
        rows = self._rows_under(gi.DataContext)
        self.model_grid.SelectedItems.Clear()
        for rv in rows:
            self.model_grid.SelectedItems.Add(rv)
        n = sum(len(self.groups[int(rv.Row['id'])].ids) for rv in rows)
        self._set_status('{} row(s), {} element(s) selected'.format(len(rows), n))
        e.Handled = True

    # -- filters -------------------------------------------------------------------
    def _model_filter(self):
        if self.model_table is None:
            return
        parts = []
        chosen = self._chosen_cats()
        if chosen is not None:
            if chosen:
                parts.append('cat IN ({})'.format(
                    ', '.join("'{}'".format(esc(c)) for c in chosen)))
            else:
                parts.append("1 = 0")
        if self.hide_noparam_cb.IsChecked:
            parts.append("hasparam = '1'")
        if self.only_empty_cb.IsChecked:
            parts.append("g7 <> '0' AND g7 <> '-'")
        if self.only_unassigned_cb.IsChecked:
            parts.append("g9 = ''")
        text = esc((self.model_search.Text or '').strip())
        if text:
            cols = ['g{}'.format(i) for i in range(len(MODEL_COLS) + len(self.extra_names))]
            parts.append('(' + ' OR '.join("{} LIKE '%{}%'".format(c, text) for c in cols) + ')')
        self.model_table.DefaultView.RowFilter = ' AND '.join(parts)
        view = self.model_table.DefaultView
        n_el = sum(int(r['g6']) for r in view)
        self.model_count.Text = '{} rows, {} elements'.format(view.Count, n_el)

    # -- category checkboxes, grouped per discipline --------------------------------------
    def _build_cat_list(self):
        before = dict((cb.Tag, cb.IsChecked) for cb in self._cat_boxes)
        counts = {}
        for g in self.groups:
            counts[g.category] = counts.get(g.category, 0) + len(g.ids)
        # every category of the project, plus anything found in the model
        names = set(counts)
        for cat in doc.Settings.Categories:
            if cat.CategoryType == DB.CategoryType.Model and cat.Name in DISCIPLINE_OF:
                names.add(cat.Name)
        per_disc = {}
        for name in names:
            per_disc.setdefault(DISCIPLINE_OF.get(name, OTHER), []).append(name)

        self.cat_list.Children.Clear()
        self._cat_boxes = []
        self._disc_boxes = []
        self._disc_name = []
        for disc in DISCIPLINES + [OTHER]:
            cats = sorted(per_disc.get(disc, []))
            if not cats:
                continue
            disc_cb = CheckBox()
            disc_cb.FontWeight = System.Windows.FontWeights.SemiBold
            disc_cb.Click += self.on_disc_check
            panel = StackPanel()
            panel.Margin = System.Windows.Thickness(22, 4, 0, 4)
            children = []
            for name in cats:
                n = counts.get(name, 0)
                cb = CheckBox()
                cb.Tag = name
                cb.Content = '{}  ({})'.format(name, n)
                cb.Margin = System.Windows.Thickness(0, 0, 0, 4)
                if not n:
                    cb.Foreground = System.Windows.Media.Brushes.Gray
                cb.IsChecked = before.get(name, True) if before else True
                cb.Click += self.on_cat_check
                panel.Children.Add(cb)
                children.append(cb)
                self._cat_boxes.append(cb)
            total = sum(counts.get(c, 0) for c in cats)
            disc_cb.Content = '{}  ({})'.format(disc, total)
            exp = Expander()
            exp.Header = disc_cb
            exp.Content = panel
            exp.Margin = System.Windows.Thickness(0, 0, 0, 4)
            self.cat_list.Children.Add(exp)
            disc_cb.Tag = children
            self._disc_boxes.append(disc_cb)
            self._disc_name.append(disc)
            self._sync_disc(disc_cb)
        self._update_cat_label()

    def _sync_disc(self, disc_cb):
        states = set(bool(cb.IsChecked) for cb in disc_cb.Tag)
        disc_cb.IsChecked = list(states)[0] if len(states) == 1 else None

    def _chosen_cats(self):
        """None = all categories (no filter), else the list of checked names."""
        chosen = [cb.Tag for cb in self._cat_boxes if cb.IsChecked]
        if len(chosen) == len(self._cat_boxes):
            return None
        return chosen

    def _update_cat_label(self):
        chosen = self._chosen_cats()
        if chosen is None:
            self.cat_label.Text = ALL_CATS
            return
        if not chosen:
            self.cat_label.Text = 'No categories'
            return
        full = [d for d in self._disc_boxes if d.IsChecked]
        covered = set(cb.Tag for d in full for cb in d.Tag)
        if full and covered == set(chosen):
            self.cat_label.Text = ', '.join(self._disc_name[i] for i, d in
                                            enumerate(self._disc_boxes) if d.IsChecked)
        elif len(chosen) == 1:
            self.cat_label.Text = chosen[0]
        else:
            self.cat_label.Text = '{} categories: {}'.format(len(chosen), ', '.join(chosen))

    def _set_all_cats(self, value):
        for cb in self._cat_boxes:
            cb.IsChecked = value
        for d in self._disc_boxes:
            d.IsChecked = value
        self._update_cat_label()

    def on_disc_check(self, sender, e):
        value = bool(sender.IsChecked)
        sender.IsChecked = value
        for cb in sender.Tag:
            cb.IsChecked = value
        self._update_cat_label()
        self._model_filter()

    def on_hide_noparam(self, sender, e):
        self._model_filter()

    def on_cat_check(self, sender, e):
        for d in self._disc_boxes:
            if sender in d.Tag:
                self._sync_disc(d)
        self._update_cat_label()
        self._model_filter()

    def on_cat_all(self, sender, e):
        self._set_all_cats(True)
        self._model_filter()

    def on_cat_none(self, sender, e):
        self._set_all_cats(False)
        self._model_filter()

    def _art_filter(self):
        if self.art_table is None:
            return
        parts = []
        sub = self.sub_cb.SelectedItem
        sub_col = self._art_col.get('OVZ_Subdiscipline')
        if sub and sub != ALL_SUBS and sub_col:
            parts.append("{} = '{}'".format(sub_col, esc(sub)))
        text = esc((self.art_search.Text or '').strip())
        if text:
            cols = list(self._art_col.values())
            parts.append('(' + ' OR '.join("{} LIKE '%{}%'".format(c, text) for c in cols) + ')')
        self.art_table.DefaultView.RowFilter = ' AND '.join(parts)
        self.art_count.Text = '{} articles'.format(self.art_table.DefaultView.Count)

    def _set_status(self, extra=''):
        with_param = [g for g in self.groups if g.has_param]
        empty = sum(g.empty for g in with_param)
        pending = sum(1 for g in self.groups if g.state in ('manual', 'saved'))
        msg = u'{}: {} element(s) empty, {} row(s) waiting to be applied'.format(
            self.target, empty, pending)
        if self.changes:
            msg += ', {} unsaved mapping change(s)'.format(len(self.changes))
        self.status_text.Text = (extra + '  |  ' if extra else '') + msg

    # -- helpers ---------------------------------------------------------------------
    def _selected_groups(self):
        return [self.groups[int(rv.Row['id'])] for rv in list(self.model_grid.SelectedItems)]

    def _selected_article(self):
        """(article code, value from the chosen column) of the selected article."""
        rv = self.art_grid.SelectedItem
        if rv is None:
            return '', ''
        code_col = self._art_col.get(ART_CODE_COL) or 'x0'
        value_col = self._art_col.get(self.value_col, code_col)
        return (rv.Row[code_col] or '').strip(), (rv.Row[value_col] or '').strip()

    def _assign(self, code, value, groups):
        if not groups:
            return
        if not value:
            self._set_status('Article {} has no value in column "{}"'.format(code, self.value_col))
            return
        for g in groups:
            g.article, g.source, g.state = value, code, 'manual'
            self._remember(g, value, code)
            # one rule per mapping key: the last value given is the one saved
            self.changes[self._mkey(g)] = self._record(g, value, code)
            self._update_row(g)
        self._set_status(u'{} row(s): {} = {}'.format(len(groups), self.target, value))

    # -- per-row values when rows share a mapping key ----------------------------------
    def _remember(self, g, value, code):
        """Store what this row got in this session ('' = removed)."""
        self._seq += 1
        for eid in g.ids:
            self._assigned[(self.target, str(eid))] = (value, code, self._seq)

    def _row_assignment(self, g):
        """Latest (value, article, order) given to this row's elements, or None."""
        best = None
        for eid in g.ids:
            a = self._assigned.get((self.target, str(eid)))
            if a is not None and (best is None or a[2] > best[2]):
                best = a
        return best

    def _latest_for_key(self, key):
        """(row, assignment) of the last value still given to a row with this key."""
        best = None
        for g in self.groups:
            if self._mkey(g) != key:
                continue
            a = self._row_assignment(g)
            if a is not None and a[0] and (best is None or a[2] > best[1][2]):
                best = (g, a)
        return best

    def _shared_key_values(self):
        """Mapping keys whose rows got different values in this session."""
        per_key = {}
        for g in self.groups:
            key = self._mkey(g)
            if not self.changes.get(key):
                continue
            a = self._row_assignment(g)
            if a is not None and a[0]:
                per_key.setdefault(key, set()).add(a[0])
        return [(key, sorted(vals)) for key, vals in sorted(per_key.items()) if len(vals) > 1]

    # -- events: columns, filters ------------------------------------------------------
    def on_auto_column(self, sender, e):
        if e.PropertyName in HIDDEN_COLS or (sender.Name == 'model_grid'
                                             and e.PropertyName in HIDDEN_MODEL_COLS):
            e.Cancel = True
            return
        # a TextBlock shows underscores literally (a plain string would hide them)
        header = TextBlock()
        header.Text = self._headers.get(e.PropertyName, e.PropertyName)
        e.Column.Header = header
        if sender.Name == 'model_grid' and e.PropertyName in COL_TIPS:
            header.ToolTip = COL_TIPS[e.PropertyName]
        if sender.Name == 'art_grid':
            e.Column.MaxWidth = 420
            if e.PropertyName == getattr(self, '_art_col', {}).get(self.value_col):
                # the column whose value is dragged: marked and tinted
                header.Text = u'\u25B6 ' + header.Text
                header.FontWeight = System.Windows.FontWeights.Bold
                style = Style(clr.GetClrType(DataGridCell),
                              self.TryFindResource(clr.GetClrType(DataGridCell)))
                style.Setters.Add(Setter(Control.BackgroundProperty,
                                         SolidColorBrush(Color.FromRgb(0xE3, 0xF4, 0xEC))))
                e.Column.CellStyle = style

    def on_model_columns_ready(self, sender, e):
        """Put the generated columns in a logical order."""
        by_name = dict((c.SortMemberPath, c) for c in self.model_grid.Columns)
        pos = 0
        for i in DISPLAY_ORDER:
            col = by_name.get('g{}'.format(i))
            if col is not None:
                col.DisplayIndex = pos
                pos += 1
        # the category is already shown in the group headers
        cat_col = by_name.get('g0')
        if cat_col is not None:
            cat_col.Visibility = Visibility.Collapsed if self._grouping() else Visibility.Visible

    def on_model_filter(self, sender, e):
        self._model_filter()

    def on_art_filter(self, sender, e):
        self._art_filter()

    # -- events: drag & drop -------------------------------------------------------------
    def on_art_mouse_down(self, sender, e):
        row = find_ancestor(e.OriginalSource, DataGridRow)
        self._drag_start = e.GetPosition(None) if row is not None else None

    def on_art_mouse_move(self, sender, e):
        if self._drag_start is None or e.LeftButton != MouseButtonState.Pressed:
            return
        pos = e.GetPosition(None)
        if (abs(pos.X - self._drag_start.X) < SystemParameters.MinimumHorizontalDragDistance
                and abs(pos.Y - self._drag_start.Y) < SystemParameters.MinimumVerticalDragDistance):
            return
        code, value = self._selected_article()
        self._drag_start = None
        if code:
            DragDrop.DoDragDrop(self.art_grid, code + '\t' + value, DragDropEffects.Copy)

    def on_drag_over(self, sender, e):
        e.Effects = DragDropEffects.Copy if e.Data.GetDataPresent(DataFormats.Text) \
            else getattr(DragDropEffects, 'None')
        e.Handled = True

    def on_drop(self, sender, e):
        data = e.Data.GetData(DataFormats.Text) or ''
        code, _, value = data.partition('\t')
        row = find_ancestor(e.OriginalSource, DataGridRow)
        gi = find_ancestor(e.OriginalSource, GroupItem) if row is None else None
        if row is not None and not row.IsSelected:
            targets = [self.groups[int(row.Item.Row['id'])]]
        elif gi is not None and isinstance(gi.DataContext, CollectionViewGroup):
            targets = self._groups_under(gi.DataContext)     # dropped on a header
        else:
            targets = self._selected_groups()
        if not targets:
            self._set_status('Select rows on the left first, or drop onto a row.')
            return
        self._assign(code, value, targets)

    def on_art_double_click(self, sender, e):
        if find_ancestor(e.OriginalSource, DataGridRow) is not None:
            self.on_assign_button(sender, e)

    def on_assign_button(self, sender, e):
        code, value = self._selected_article()
        targets = self._selected_groups()
        if not code or not targets:
            self._set_status('Select rows on the left and an article on the right.')
            return
        self._assign(code, value, targets)

    # -- running things inside Revit ----------------------------------------------------------
    def _run(self, func):
        """Ask Revit to run func() as soon as it is ready."""
        if self._busy:
            self._set_status('Still busy with the previous action...')
            return
        self._busy = True

        def wrapped(uiapp):
            try:
                if not doc.IsValidObject:
                    forms.alert('The model this window was opened for is closed. '
                                'Close the window and start the tool again.')
                    return
                active = uiapp.ActiveUIDocument
                if active is None or not active.Document.Equals(doc):
                    forms.alert('Switch back to "{}" in Revit first.'.format(doc.Title))
                    return
                func()
            except Exception as ex:
                forms.alert('Stopped: {}'.format(ex))
            finally:
                self._busy = False
        ACTION.func = wrapped
        EVENT.Raise()
        self.status_text.Text = 'Working in Revit...'

    def on_refresh(self, sender, e):
        def run():
            self.load_model()
            self._set_status('Model reloaded')
        self._run(run)

    # -- events: model actions -------------------------------------------------------------
    def _ids(self, groups):
        ids = List[DB.ElementId]()
        for g in groups:
            for i in g.ids:
                ids.Add(i)
        return ids

    def on_select(self, sender, e):
        groups = self._selected_groups()
        if not groups:
            return
        ids = self._ids(groups)

        def run():
            self._own_selection = set(str(i) for i in ids)
            uidoc.Selection.SetElementIds(ids)
            uidoc.RefreshActiveView()
            self._set_status('{} element(s) selected in Revit'.format(ids.Count))
        self._run(run)

    def on_isolate(self, sender, e):
        groups = self._selected_groups()
        if not groups:
            return
        ids = self._ids(groups)

        def run():
            view = doc.ActiveView
            t = DB.Transaction(doc, 'OVZ Param Mapper: isolate')
            try:
                t.Start()
                view.IsolateElementsTemporary(ids)
                t.Commit()
            except Exception as ex:
                if t.HasStarted() and not t.HasEnded():
                    t.RollBack()
                forms.alert('Could not isolate in this view: {}'.format(ex))
                return
            uidoc.RefreshActiveView()
            self._set_status('Isolated in "{}" (reset with the sunglasses icon in Revit)'
                             .format(view.Name))
        self._run(run)

    def on_clear(self, sender, e):
        groups = [g for g in self._selected_groups() if g.article]
        for g in groups:
            self._remember(g, '', '')
            g.article, g.source, g.state = '', '', ''
            self._update_row(g)
        for key in set(self._mkey(g) for g in groups):
            # another row with the same key still has a value: keep that as the rule
            left = self._latest_for_key(key)
            if left is not None:
                self.changes[key] = self._record(left[0], left[1][0], left[1][1])
            elif key in self.mapping or key in self.changes:
                self.changes[key] = None
        if groups:
            self._set_status('Value removed from {} row(s); the model keeps its value '
                             'until you change it'.format(len(groups)))

    # -- check colours in a 3D view -------------------------------------------------------
    def _check_view(self, create=True):
        """The personal 'OVZ check - <user>' 3D view (made from the active 3D view
        or as a new isometric view)."""
        name = u'OVZ check - {}'.format(USER or 'user')
        for v in DB.FilteredElementCollector(doc).OfClass(DB.View3D):
            if not v.IsTemplate and v.Name == name:
                return v
        if not create:
            return None
        active = doc.ActiveView
        view = None
        if isinstance(active, DB.View3D) and not active.IsTemplate and \
                active.CanViewBeDuplicated(DB.ViewDuplicateOption.Duplicate):
            view = doc.GetElement(active.Duplicate(DB.ViewDuplicateOption.Duplicate))
        else:
            vft = [t for t in DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType)
                   if t.ViewFamily == DB.ViewFamily.ThreeDimensional]
            view = DB.View3D.CreateIsometric(doc, vft[0].Id)
        try:
            view.ViewTemplateId = DB.ElementId.InvalidElementId   # overrides must be allowed
        except Exception:
            pass
        view.Name = name
        return view

    def _check_status(self, g, el, pname, on_type):
        """'empty', 'different', 'ok' or None (no mapping to compare with)."""
        target = doc.GetElement(el.GetTypeId()) if on_type else el
        p = target.LookupParameter(pname) if target is not None else None
        if p is None:
            return None
        value = read_value(p)
        if not value:
            return 'empty'
        if not g.article:
            return None
        if value == g.article:
            return 'ok'
        # numbers: '300 mm' = '300' (never for text, where 68.11.10.250 != 68.11.10.710)
        if p.StorageType != DB.StorageType.String and _same_number(value, g.article):
            return 'ok'
        return 'different'

    def on_colour(self, sender, e):
        rows = [self.groups[int(rv.Row['id'])] for rv in self.model_table.DefaultView]
        rows = [g for g in rows if g.has_param]
        if not rows:
            forms.alert('No rows in the list to check. Change the filters on the left.')
            return
        with_green = bool(self.green_cb.IsChecked)
        pname, on_type = split_target(self.target)

        def run():
            solid = None
            for fp in DB.FilteredElementCollector(doc).OfClass(DB.FillPatternElement):
                if fp.GetFillPattern().IsSolidFill:
                    solid = fp.Id
                    break
            colours = {'empty': DB.Color(239, 65, 54), 'different': DB.Color(245, 150, 20),
                       'ok': DB.Color(40, 170, 90)}
            settings = {}
            for key, c in colours.items():
                ogs = DB.OverrideGraphicSettings()
                ogs.SetProjectionLineColor(c)
                ogs.SetCutLineColor(c)
                if solid is not None:
                    ogs.SetSurfaceForegroundPatternId(solid)
                    ogs.SetSurfaceForegroundPatternColor(c)
                    ogs.SetCutForegroundPatternId(solid)
                    ogs.SetCutForegroundPatternColor(c)
                settings[key] = ogs
            clear = DB.OverrideGraphicSettings()
            counts = {'empty': 0, 'different': 0, 'ok': 0}
            t = DB.Transaction(doc, 'OVZ Param Mapper: check colours')
            try:
                t.Start()
                view = self._check_view()
                for g in self.groups:                 # remove old colours first
                    for eid in g.ids:
                        view.SetElementOverrides(eid, clear)
                for g in rows:
                    for eid in g.ids:
                        el = doc.GetElement(eid)
                        if el is None:
                            continue
                        state = self._check_status(g, el, pname, on_type)
                        if state is None or (state == 'ok' and not with_green):
                            continue
                        view.SetElementOverrides(eid, settings[state])
                        counts[state] += 1
                t.Commit()
            except Exception as ex:
                if t.HasStarted() and not t.HasEnded():
                    t.RollBack()
                forms.alert('Could not colour the elements: {}'.format(ex))
                return
            uidoc.RequestViewChange(view)
            msg = u'"{}": {} red (empty), {} orange (different)'.format(
                view.Name, counts['empty'], counts['different'])
            if with_green:
                msg += u', {} green (correct)'.format(counts['ok'])
            self._set_status(msg)
        self._run(run)

    def on_uncolour(self, sender, e):
        def run():
            view = self._check_view(create=False)
            if view is None:
                self._set_status('There is no check view to clear')
                return
            clear = DB.OverrideGraphicSettings()
            t = DB.Transaction(doc, 'OVZ Param Mapper: clear colours')
            try:
                t.Start()
                for g in self.groups:
                    for eid in g.ids:
                        view.SetElementOverrides(eid, clear)
                t.Commit()
            except Exception as ex:
                if t.HasStarted() and not t.HasEnded():
                    t.RollBack()
                forms.alert('Could not clear the colours: {}'.format(ex))
                return
            uidoc.RefreshActiveView()
            self._set_status(u'Colours cleared in "{}"'.format(view.Name))
        self._run(run)

    def on_unfilled(self, sender, e):
        self._set_all_cats(True)
        self.model_search.Text = ''
        self.only_unassigned_cb.IsChecked = False
        self.only_empty_cb.IsChecked = True
        self.hide_noparam_cb.IsChecked = True
        self._model_filter()
        empty = sum(g.empty for g in self.groups if g.has_param)
        rows = sum(1 for g in self.groups if g.has_param and g.empty)
        self._set_status(u'Find unfilled: {} element(s) in {} row(s) have no {}'
                         .format(empty, rows, self.target))

    # -- events: files -----------------------------------------------------------------------
    def on_list_browse(self, sender, e):
        path = forms.pick_file(file_ext='xlsx')
        if path:
            self.load_articles(path)

    def on_map_open(self, sender, e):
        if self.changes and not self._ask_save():
            return
        path = forms.pick_file(file_ext='csv')
        if path:
            self.load_mapping_file(path)

    def on_map_new(self, sender, e):
        if self.changes and not self._ask_save():
            return
        path = forms.save_file(file_ext='csv', default_name='OVZ_ParamMapper_mapping')
        if not path:
            return
        write_mapping(path, {})
        self.load_mapping_file(path)

    def save_mapping(self):
        """Merge our changes into the shared file. Returns False if cancelled."""
        if not self.changes:
            return True
        if not self.map_path:
            path = forms.save_file(file_ext='csv', default_name='OVZ_ParamMapper_mapping')
            if not path:
                return False
            self.map_path = path
        shared = self._shared_key_values()
        if shared:
            lines = [u'{}: {}   (saved: {})'.format(key.replace('|', ' / '), ', '.join(vals),
                                                  self.changes[key]['Value'])
                     for key, vals in shared[:12]]
            more = '\n...and {} more'.format(len(shared) - 12) if len(shared) > 12 else ''
            forms.alert(u'These rows share one rule in the shared mapping, so only one value '
                        u'can be remembered for the next project:\n\n{}{}\n\nThe last value is '
                        u'saved. In this model, every row is written with its own value.'
                        .format('\n'.join(lines), more))
        try:
            disk = load_mapping(self.map_path)
        except Exception as ex:
            forms.alert('Could not read the mapping file: {}'.format(ex))
            return False
        conflicts = find_conflicts(disk, self.snapshot, self.changes)
        skip = []
        if conflicts:
            lines = []
            for key, theirs, mine in conflicts[:12]:
                who = theirs['Updated by'] if theirs else 'someone'
                lines.append(u'{}\n   theirs ({}): {}   yours: {}'.format(
                    key.replace('|', ' / '), who,
                    theirs['Value'] if theirs else '(removed)',
                    mine['Value'] if mine else '(removed)'))
            more = '\n...and {} more'.format(len(conflicts) - 12) if len(conflicts) > 12 else ''
            choice = forms.alert(
                u'{} rule(s) were changed by a colleague since you opened the mapping:\n\n'
                u'{}{}\n\nWhich version should be kept?'.format(
                    len(conflicts), '\n'.join(lines), more),
                options=['Keep mine', 'Keep theirs', 'Cancel'])
            if not choice or choice == 'Cancel':
                return False
            if choice == 'Keep theirs':
                skip = [c[0] for c in conflicts]
        merged = merge(disk, self.changes, skip)
        try:
            write_mapping(self.map_path, merged)
        except Exception as ex:
            forms.alert('Could not save the mapping file (is it open in Excel?)\n\n{}'.format(ex))
            return False
        self.mapping = merged
        self.snapshot = dict(merged)
        self.changes = {}
        self.map_box.Text = '{}  ({} rules)'.format(self.map_path, len(merged))
        config.ovz_mapping = self.map_path
        script.save_config()
        self._apply_mapping_to_groups()
        return True

    def _ask_save(self):
        choice = forms.alert('You have {} unsaved mapping change(s). Save them to the '
                             'shared mapping first?'.format(len(self.changes)),
                             options=['Save', "Don't save", 'Cancel'])
        if choice == 'Save':
            return self.save_mapping()
        if choice == "Don't save":
            self.changes = {}
            return True
        return False

    def on_save_as(self, sender, e):
        """Write the current rules + your changes to a NEW file and continue there.
        The original file is not touched."""
        path = forms.save_file(file_ext='csv', default_name='OVZ_ParamMapper_mapping_copy')
        if not path:
            return
        base = dict(self.mapping)
        if self.map_path and os.path.exists(self.map_path):
            try:
                base = load_mapping(self.map_path)
            except Exception:
                pass
        merged = merge(base, self.changes)
        try:
            write_mapping(path, merged)
        except Exception as ex:
            forms.alert('Could not save the file.\n\n{}'.format(ex))
            return
        self.map_path = path
        self.mapping = merged
        self.snapshot = dict(merged)
        self.changes = {}
        self.map_box.Text = '{}  ({} rules)'.format(path, len(merged))
        config.ovz_mapping = path
        script.save_config()
        self._apply_mapping_to_groups()
        self._set_status('Mapping saved as a new file; you are now working in that file')

    def on_save(self, sender, e):
        if self.save_mapping():
            self._set_status('Mapping saved')

    # -- events: apply ------------------------------------------------------------------------
    def on_apply(self, sender, e):
        todo = [g for g in self.groups if g.article and g.has_param
                and (g.current != g.article)]
        if not todo:
            forms.alert('Nothing to write: every row with a new value already has it.')
            return
        pname, on_type = split_target(self.target)
        if on_type:
            # a type parameter holds ONE value per type: check rows that share a type
            per_type = {}
            for g in todo:
                if g.type_id is not None:
                    per_type.setdefault(str(g.type_id), set()).add(g.article)
            clash = [k for k, v in per_type.items() if len(v) > 1]
            if clash:
                if not forms.alert(
                        '"{}" is a type parameter, but {} type(s) got different values in '
                        'different rows (e.g. per size). Those types will be skipped.\n\n'
                        'Continue with the rest?'.format(pname, len(clash)), yes=True, no=True):
                    return
                todo = [g for g in todo if str(g.type_id) not in clash]
            n_types = len(set(str(g.type_id) for g in todo))
            question = 'Write "{}" on {} type(s)?'.format(pname, n_types)
        else:
            n_el = sum(len(g.ids) for g in todo)
            question = 'Write "{}" on {} element(s) in {} row(s)?'.format(pname, n_el, len(todo))
        if not todo or not forms.alert(question + '\nYou can undo this with Ctrl+Z in Revit.',
                                       yes=True, no=True):
            return
        self._run(lambda: self._write(todo))

    def _write(self, todo):
        written = 0
        pname, on_type = split_target(self.target)
        problems = {}
        t = DB.Transaction(doc, 'OVZ Param Mapper: {}'.format(pname))
        try:
            t.Start()
            done_types = set()
            for g in todo:
                if on_type:
                    targets = [doc.GetElement(g.type_id)] if g.type_id is not None else []
                    if str(g.type_id) in done_types:
                        targets = []
                    done_types.add(str(g.type_id))
                else:
                    targets = [doc.GetElement(eid) for eid in g.ids]
                ok_all = True
                shown = None
                for el in targets:
                    p = el.LookupParameter(pname) if el is not None else None
                    if p is None or p.IsReadOnly:
                        continue
                    try:
                        if write_value(p, g.article):
                            written += 1
                        shown = read_value(p)
                    except Exception as ex:
                        ok_all = False
                        problems[str(ex)] = problems.get(str(ex), 0) + 1
                        continue
                if ok_all and shown is not None:
                    g.values = set([shown])
                    g.empty = 0 if shown else len(g.ids)
            if on_type:
                # other rows of the same types now show the new value too
                for g in self.groups:
                    if g.type_id is not None and str(g.type_id) in done_types and g not in todo:
                        p = doc.GetElement(g.type_id).LookupParameter(pname)
                        if p is not None:
                            v = read_value(p)
                            g.values = set([v])
                            g.empty = 0 if v else len(g.ids)
                            self._update_row(g)
            t.Commit()
        except Exception as ex:
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
            forms.alert('Stopped, nothing was changed: {}'.format(ex))
            return
        for g in todo:
            g.state = 'done'
            self._update_row(g)
        if problems:
            forms.alert(u'Some values could not be written into "{}":\n\n{}'.format(
                pname, '\n'.join(u'{} ({} times)'.format(k, v) for k, v in problems.items())))
        saved = self.save_mapping()
        self._model_filter()
        self._set_status('"{}" written {} time(s){}'.format(
            pname, written, '' if saved else ' (mapping NOT saved)'))

    # -- closing --------------------------------------------------------------------------------
    def on_close_click(self, sender, e):
        self.Close()

    def on_closing(self, sender, e):
        if self.changes and not self._closing_ok:
            if not self._ask_save():
                e.Cancel = True
                return
        self._closing_ok = True
        self._unsubscribe_selection()

    # -- follow the selection in Revit ------------------------------------------------------
    def _subscribe_selection(self):
        """Revit 2023+ tells us when the selection changes in the model."""
        try:
            self._sel_callback = self.on_revit_selection
            HOST_APP.uiapp.SelectionChanged += self._sel_callback
        except Exception:
            self._sel_callback = None
            self.follow_cb.IsEnabled = False
            self.follow_cb.ToolTip = 'Needs Revit 2023 or newer'

    def _unsubscribe_selection(self):
        if self._sel_callback is not None:
            try:
                HOST_APP.uiapp.SelectionChanged -= self._sel_callback
            except Exception:
                pass
            self._sel_callback = None

    def on_follow_toggle(self, sender, e):
        config.ovz_follow_selection = bool(self.follow_cb.IsChecked)
        script.save_config()

    def on_revit_selection(self, sender, args):
        if not self.follow_cb.IsChecked or self.model_table is None:
            return
        try:
            if not args.GetDocument().Equals(doc):
                return
            ids = [str(i) for i in args.GetSelectedElements()]
        except Exception:
            return
        if not ids:
            return
        if set(ids) == self._own_selection:      # we selected these ourselves
            self._own_selection = set()
            return
        wanted = set()
        for i in ids:
            idx = self._el_to_group.get(i)
            if idx is not None:
                wanted.add(str(idx))
        if not wanted:
            self._set_status('Selected in Revit: not in this list (other category, or '
                             'click Refresh model if it is new)')
            return
        rows = [rv for rv in self.model_table.DefaultView if rv.Row['id'] in wanted]
        self.model_grid.SelectedItems.Clear()
        for rv in rows:
            self.model_grid.SelectedItems.Add(rv)
        if rows:
            self.model_grid.ScrollIntoView(rows[0])
        msg = u'Selected in Revit: {} element(s) \u2192 {} row(s)'.format(len(ids), len(rows))
        hidden = len(wanted) - len(rows)
        if hidden:
            msg += u', {} row(s) hidden by the current filters'.format(hidden)
        self._set_status(msg)


if doc is None:
    forms.alert('Open a Revit model first.', exitscript=True)
if doc.IsFamilyDocument:
    forms.alert('OVZ Param Mapper works in project files, not in the Family Editor.',
                exitscript=True)

window = OvzWindow()
try:
    # lets text boxes receive typing while the window stays open next to Revit
    clr.AddReference('WindowsFormsIntegration')
    from System.Windows.Forms.Integration import ElementHost
    ElementHost.EnableModelessKeyboardInterop(window)
except Exception:
    pass
window.Show()
