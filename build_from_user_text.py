import os
import re
import subprocess
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

TECHNICAL_WORDS = {
    'superoptimization', 'Superoptimization', 'superoptimizer', 'Superoptimizers',
    'calldata', 'Calldata', 'solc', 'sload', 'sstore', 'mstore', 'mload', 'mcopy',
    'tload', 'tstore', 'TLOAD', 'TSTORE', 'SLOAD', 'SSTORE', 'MLOAD', 'MSTORE', 'MCOPY',
    'revm', 'evmone', 'OpenZeppelin', 'ReentrancyGuard', 'PackedSlot', 'Unhoisted',
    'stateVar', 'dest_ptr', 'src_ptr', 'zkEVMs', 'zkEVM', 'arithmetization',
    'jumpdest', 'Gwei', 'cancun', 'Cancun', 'bitvector', 'EthIR', 'MadMax', 'Securify',
    'ebso', 'Ebso', 'I_o', 'I_v', 'I_d', 'I_b', 'o_k', 'DUPk', 'SWAPk', 'ISZERO', 'JUMPDEST',
    'PUSH0', 'PUSH1', 'PUSH32', 'DUP1', 'DUP16', 'SWAP1', 'SWAP16', 'CALLDATALOAD',
    'CALLDATASIZE', 'DELEGATECALL', 'STATICCALL', 'REVERT', 'try/finally', 'runs',
    'addr', 'accum', 'init', 'uint128', 'uint64', 'uint256', 'uint8', 'uint32',
    'secp256k1', 'Keccak-256', 'vvvv', 'mainnet', 'Massalin', 'Schett', 'lowerer',
    'lowerers', 'trie', 'struct', 'SlotIndex', 'BitOffset', 'DomTree', 'FreshReg', 'inliner',
    'ir', 'evm', 'QF_BV', 'Z3', 'CVC5', 'via-ir', 'gasleft', 'Vyper',
    'RJUMP', 'RJUMPI', 'RJUMPV', 'CALLF', 'RETF', 'adc', 'opcode', 'opcodes',
    'DUP1-DUP16', 'SWAP1-SWAP16', 'YUL', "EVM's", 'Post-Cancun', 'ISA', 'EIP-170',
    'Erigon', 'Reth', 'CEGIS', 'precompile', 'precompiles', 'Precompile', 'Precompiles',
    'bn256Add', 'ecrecover', 'modexp', 'DAG', 'DAGs', 'NP-hard', 'bytecode', 'bytecodes',
    'dispatch', 'dispatcher', 'dispatchers', 'canonicalization', 'unhoisted', 'multicall',
    'Multicall', 'polyhedral', 'abi', 'ABI', 'calldatasize', 'calldataload', 'allocator',
    'allocators', 'msg.sig', 'JUMPI', 'EQ', 'decoupling', 'pre-warming', 'pre-warm',
    'preconditions', 'reentrancy', 'Reentrancy', 'namespaces', 'unaligned',
    'https://github.com/Risheek15/EVM-Compiler-Optimization', 'EVM-Compiler-Optimization',
    'Risheek', 'Teja', 'Vellore', 'taxanomy', 'superoptimisation', 'Superoptimisation',
    'optimisations', 'Optimisations', 'optimised', 'Optimised', 'optimising', 'Optimising',
    'formalises', 'Formalises', 'decentralised', 'Decentralised', 'modernised', 'Modernised',
    'analyse', 'Analyse', 'analysed', 'Analysed', 'modelling', 'Modelling', 'QF_BV'
}

def mark_no_proof(run):
    rPr = run._r.get_or_add_rPr()
    rPr.append(parse_xml(f'<w:noProof {nsdecls("w")}/>'))

def add_smart_p(doc, text, indent=True, font_size=9.5, line_spacing=1.0, space_before=0, space_after=1.5):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing
    if indent:
        p.paragraph_format.first_line_indent = Inches(0.15)

    tech_lower = {tw.lower() for tw in TECHNICAL_WORDS}
    words = text.split(' ')
    for i, w in enumerate(words):
        clean_w = re.sub(r'^[^\w]+|[^\w]+$', '', w)
        r = p.add_run(w + (' ' if i < len(words) - 1 else ''))
        r.font.name = 'Times New Roman'
        r.font.size = Pt(font_size)
        if clean_w.lower() in tech_lower or any(clean_w.startswith(prefix) for prefix in ['0x', 'EIP-', 'ERC-', 'g_', 'C_', 'G_', 'R_', 'B_', 'I_', 'W_']):
            mark_no_proof(r)
    return p

def add_bullet_p(doc, text, font_size=9.5, line_spacing=1.0, space_before=0, space_after=1.5):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing
    p.paragraph_format.left_indent = Inches(0.20)
    p.paragraph_format.first_line_indent = Inches(-0.10)

    tech_lower = {tw.lower() for tw in TECHNICAL_WORDS}
    words = text.split(' ')
    for i, w in enumerate(words):
        clean_w = re.sub(r'^[^\w]+|[^\w]+$', '', w)
        r = p.add_run(w + (' ' if i < len(words) - 1 else ''))
        r.font.name = 'Times New Roman'
        r.font.size = Pt(font_size)
        if clean_w.lower() in tech_lower or any(clean_w.startswith(prefix) for prefix in ['0x', 'EIP-', 'ERC-', 'g_', 'C_', 'G_', 'R_', 'B_', 'I_', 'W_']):
            mark_no_proof(r)
    return p

def create_ieee_table(doc, headers, rows, col_widths, caption):
    p_cap = doc.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.paragraph_format.space_before = Pt(4.5)
    p_cap.paragraph_format.space_after = Pt(1.5)
    p_cap.paragraph_format.keep_with_next = True
    r_cap = p_cap.add_run(caption.upper())
    r_cap.font.name = 'Times New Roman'
    r_cap.font.size = Pt(7.5)
    r_cap.font.bold = True

    total_dxa = sum(int(w * 1440) for w in col_widths)

    tbl = doc.add_table(rows=len(rows)+1, cols=len(headers))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False

    tblPr = tbl._tbl.tblPr
    for child in list(tblPr):
        if child.tag.endswith('tblW') or child.tag.endswith('tblBorders'):
            tblPr.remove(child)

    tblW = parse_xml(f'<w:tblW {nsdecls("w")} w:w="{total_dxa}" w:type="dxa"/>')
    tblPr.append(tblW)

    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="8" w:space="0" w:color="000000"/>'
        f'<w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/>'
        f'<w:left w:val="none"/>'
        f'<w:right w:val="none"/>'
        f'<w:insideH w:val="none"/>'
        f'<w:insideV w:val="none"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

    hdr_row = tbl.rows[0]
    trPr = hdr_row._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

    for i, h_text in enumerate(headers):
        cell = hdr_row.cells[i]
        cell.width = Inches(col_widths[i])
        tcPr = cell._tc.get_or_add_tcPr()
        for child in list(tcPr):
            if child.tag.endswith('tcW'):
                tcPr.remove(child)
        tcW = parse_xml(f'<w:tcW {nsdecls("w")} w:w="{int(col_widths[i]*1440)}" w:type="dxa"/>')
        tcPr.append(tcW)
        
        tcBorders = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            f'<w:top w:val="none"/><w:left w:val="none"/><w:right w:val="none"/>'
            f'</w:tcBorders>'
        )
        tcPr.append(tcBorders)

        tcMar = parse_xml(
            f'<w:tcMar {nsdecls("w")}>'
            f'<w:top w:w="30" w:type="dxa"/><w:bottom w:w="30" w:type="dxa"/>'
            f'<w:left w:w="40" w:type="dxa"/><w:right w:w="40" w:type="dxa"/>'
            f'</w:tcMar>'
        )
        tcPr.append(tcMar)

        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.0
        run = p.add_run(h_text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(7.0)
        run.font.bold = True
        mark_no_proof(run)

    for r_idx, r_data in enumerate(rows):
        row = tbl.rows[r_idx+1]
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

        for c_idx, val in enumerate(r_data):
            cell = row.cells[c_idx]
            cell.width = Inches(col_widths[c_idx])
            tcPr = cell._tc.get_or_add_tcPr()
            for child in list(tcPr):
                if child.tag.endswith('tcW'):
                    tcPr.remove(child)
            tcW = parse_xml(f'<w:tcW {nsdecls("w")} w:w="{int(col_widths[c_idx]*1440)}" w:type="dxa"/>')
            tcPr.append(tcW)

            tcMar = parse_xml(
                f'<w:tcMar {nsdecls("w")}>'
                f'<w:top w:w="20" w:type="dxa"/><w:bottom w:w="20" w:type="dxa"/>'
                f'<w:left w:w="40" w:type="dxa"/><w:right w:w="40" w:type="dxa"/>'
                f'</w:tcMar>'
            )
            tcPr.append(tcMar)

            p = cell.paragraphs[0]
            if c_idx == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p.paragraph_format.space_before = Pt(0.5)
            p.paragraph_format.space_after = Pt(0.5)
            p.paragraph_format.line_spacing = 1.0
            run = p.add_run(str(val))
            run.font.name = 'Times New Roman'
            run.font.size = Pt(7.0)
            mark_no_proof(run)

    p_after = doc.add_paragraph()
    p_after.paragraph_format.space_before = Pt(0.5)
    p_after.paragraph_format.space_after = Pt(1.5)

def build_document(text_path, output_path, h1_gap=12.0):
    with open(text_path, 'r', encoding='utf-8') as f:
        full_text = f.read()

    lines = [l.strip() for l in full_text.splitlines() if l.strip()]

    doc = Document()

    # Section 1: Title and Author (1 Column)
    sec1 = doc.sections[0]
    sec1.top_margin = Inches(0.75)
    sec1.bottom_margin = Inches(0.75)
    sec1.left_margin = Inches(0.625)
    sec1.right_margin = Inches(0.625)
    sec1.page_width = Inches(8.5)
    sec1.page_height = Inches(11.0)

    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(10)
    normal_style.font.color.rgb = RGBColor(0, 0, 0)
    normal_style.paragraph_format.line_spacing = 1.0
    normal_style.paragraph_format.space_after = Pt(0)
    normal_style.paragraph_format.space_before = Pt(0)

    # Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(8)
    run_title = p_title.add_run(lines[0])
    run_title.font.name = 'Times New Roman'
    run_title.font.size = Pt(20)
    run_title.font.bold = True

    # Author
    p_author = doc.add_paragraph()
    p_author.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_author.paragraph_format.space_after = Pt(2)
    run_author = p_author.add_run(lines[1])
    run_author.font.name = 'Times New Roman'
    run_author.font.size = Pt(11.5)
    run_author.font.bold = True

    # Affiliation
    p_affil = doc.add_paragraph()
    p_affil.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_affil.paragraph_format.space_after = Pt(10)
    affil_text = f"{lines[2]}\n{lines[3]}\n{lines[4]}"
    run_affil = p_affil.add_run(affil_text)
    run_affil.font.name = 'Times New Roman'
    run_affil.font.size = Pt(9.5)
    run_affil.font.italic = True
    mark_no_proof(run_affil)

    # Section 2: 2 columns
    sec2 = doc.add_section(docx.enum.section.WD_SECTION.CONTINUOUS)
    sec2.top_margin = Inches(0.75)
    sec2.bottom_margin = Inches(0.75)
    sec2.left_margin = Inches(0.625)
    sec2.right_margin = Inches(0.625)

    sectPr = sec2._sectPr
    cols = sectPr.xpath('./w:cols')
    if cols:
        cols[0].set(qn('w:num'), '2')
        cols[0].set(qn('w:space'), '360')
    else:
        new_cols = OxmlElement('w:cols')
        new_cols.set(qn('w:num'), '2')
        new_cols.set(qn('w:space'), '360')
        sectPr.append(new_cols)

    def add_abstract(abstract_text, index_terms_text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.0
        
        r_lead = p.add_run("Abstract — ")
        r_lead.font.name = 'Times New Roman'
        r_lead.font.size = Pt(9)
        r_lead.font.bold = True
        r_lead.font.italic = True
        
        # clean lead from text
        body = re.sub(r'^Abstract\s*[\—\-\–]\s*', '', abstract_text)
        r_body = p.add_run(body)
        r_body.font.name = 'Times New Roman'
        r_body.font.size = Pt(9)
        r_body.font.bold = True

        p_idx = doc.add_paragraph()
        p_idx.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_idx.paragraph_format.space_before = Pt(1)
        p_idx.paragraph_format.space_after = Pt(6)
        p_idx.paragraph_format.line_spacing = 1.0

        r_idx_lead = p_idx.add_run("Keywords — ")
        r_idx_lead.font.name = 'Times New Roman'
        r_idx_lead.font.size = Pt(9)
        r_idx_lead.font.bold = True
        r_idx_lead.font.italic = True

        idx_body = re.sub(r'^Keywords\s*[\—\-\–]\s*', '', index_terms_text)
        r_idx_body = p_idx.add_run(idx_body)
        r_idx_body.font.name = 'Times New Roman'
        r_idx_body.font.size = Pt(9)
        r_idx_body.font.bold = True

    def add_h1(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(h1_gap)
        p.paragraph_format.space_after = Pt(h1_gap)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text.upper())
        r.font.name = 'Times New Roman'
        r.font.size = Pt(9.5)
        r.font.bold = True
        mark_no_proof(r)
        return p

    def add_h2(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(7.5)
        p.paragraph_format.space_after = Pt(4.5)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(9.5)
        r.font.italic = True
        r.font.bold = True
        mark_no_proof(r)
        return p

    def add_h3(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(4.0)
        p.paragraph_format.space_after = Pt(2.0)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(9.5)
        r.font.italic = True
        mark_no_proof(r)
        return p

    def add_eq(eq_line):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(1.5)
        p.paragraph_format.space_after = Pt(1.5)
        p.paragraph_format.left_indent = Inches(0.12)
        
        m = re.search(r'(\(\d+\))$', eq_line)
        if m:
            eq_num = m.group(1)
            eq_text = eq_line[:m.start()].strip()
        else:
            eq_text = eq_line
            eq_num = ""

        r_eq = p.add_run(eq_text)
        r_eq.font.name = 'Cambria Math'
        r_eq.font.size = Pt(9)
        r_eq.font.italic = True
        mark_no_proof(r_eq)
        
        if eq_num:
            r_tab = p.add_run(f"    \t{eq_num}")
            r_tab.font.name = 'Times New Roman'
            r_tab.font.size = Pt(9)
            r_tab.font.italic = False
            mark_no_proof(r_tab)
        return p

    def add_code(code_text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(1.5)
        p.paragraph_format.space_after = Pt(1.5)
        p.paragraph_format.left_indent = Inches(0.04)
        p.paragraph_format.right_indent = Inches(0.04)
        
        pPr = p._p.get_or_add_pPr()
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F4F4F6"/>')
        pPr.append(shd)
        
        r = p.add_run(code_text)
        r.font.name = 'Consolas'
        r.font.size = Pt(6.8)
        mark_no_proof(r)
        return p

    # Abstract & Keywords
    add_abstract(lines[5], lines[6])

    # Table definitions
    headers_t1 = ['Opcode/Context', 'Cost (Gas)', 'Reference']
    rows_t1 = [
        ['ADD, SUB, NOT, LT, GT', '3', 'Yellow Paper [1]'],
        ['MUL, DIV, SDIV, MOD', '5', 'Yellow Paper [1]'],
        ['EXP', '10 + 50 ceil(log)', 'Yellow Paper [1]'],
        ['PUSH0', '2', 'EIP-3855 [13]'],
        ['PUSH1-PUSH32', '3', 'Yellow Paper [1]'],
        ['DUP1-16, SWAP1-16', '3', 'Yellow Paper [1]'],
        ['MLOAD, MSTORE (Warm)', '3', 'Yellow Paper [1]'],
        ['MCOPY (Word Aligned)', '3 + 3 ceil(w)', 'EIP-5656 [14]'],
        ['TLOAD/TSTORE', '100', 'EIP-1153 [4]'],
        ['SLOAD (Warm)', '100', 'EIP-2929 [5]'],
        ['SLOAD (Cold)', '2100', 'EIP-2929 [5]'],
        ['SSTORE (0->Non-Zero, Cold)', '22100', 'EIP-2929/3529'],
        ['SSTORE (Non-0->Non-0, Cold)', '5000', 'EIP-2929/3529'],
        ['SSTORE (Dirty Slot)', '100', 'EIP-3529 [6]'],
        ['CALLDATA/CALLDATASIZE', '3/2', 'Yellow Paper [1]'],
        ['JUMPDEST', '1', 'Yellow Paper [1]']
    ]

    headers_t2 = ['Pass Name', 'Scope', 'Description']
    rows_t2 = [
        ['Disambiguator', 'Module', 'All names have unique identifiers'],
        ['Function Inliner', 'Call Graph', 'Inlining if threshold reached'],
        ['Expr. Simplifier', 'Basic Block', 'Algebraic term rewriting'],
        ['CSE', 'SSA Block', 'Removing duplicate expressions'],
        ['Rematerialiser', 'Local', 'Recalculate cheap literals'],
        ['Dead Code Elim.', 'CFG', 'Remove unreachable blocks & variables'],
        ['LICM', 'Loop Header', 'Move invariant expressions'],
        ['Forwarder', 'Dominator', 'Replace SLOAD/MLOAD with value'],
        ['Stack Compressor', 'Function Frame', 'Reorder variables to satisfy 16 depth']
    ]

    headers_t3 = ['Benchmark Pattern', 'Baseline Gas', 'Optimized Gas', 'Absolute Savings', 'Relative Savings']
    rows_t3 = [
        ['Reentrancy Lock', '5,100', '303', '4,797', '94.06%'],
        ['Slot Packing', '66,300', '22,524', '43,776', '66.03%'],
        ['Loop Invariant (N=100)', '12,000', '2,400', '9,600', '80.00%'],
        ['Memory Copy (256B)', '712', '57', '655', '91.99%'],
        ['Custom Error', '147', '24', '123', '83.67%']
    ]

    headers_t4 = ['Call Depth / Context', 'State Access Type', 'Cold Cost', 'Warm Cost']
    rows_t4 = [
        ['Root Frame', 'Unaccessed (Cold)', '2100 gas', '100 gas'],
        ['Internal CALL', 'Warm', '100 gas', '100 gas'],
        ['STATICCALL', 'Transient Storage', '100 gas', '100 gas'],
        ['DELEGATECALL', 'Dirty', '100 gas', '100 gas']
    ]

    # Process remaining lines starting at index 7
    i = 7
    in_table = False
    in_algo = False
    algo_lines = []

    while i < len(lines):
        line = lines[i]

        # Check H1
        if re.match(r'^(?:I|II|III|IV|V|VI|VII|VIII|IX)\.\s+[A-Z\s&,]+$', line) or line in ['REPRODUCIBILITY', 'REFERENCES']:
            if in_algo and algo_lines:
                add_code('\n'.join(algo_lines))
                algo_lines = []
                in_algo = False
            in_table = False
            add_h1(line)
            i += 1
            continue

        # Check Tables
        if line.startswith('TABLE I:'):
            if in_algo and algo_lines:
                add_code('\n'.join(algo_lines))
                algo_lines = []
                in_algo = False
            create_ieee_table(doc, headers_t1, rows_t1, [1.35, 0.95, 1.10], line)
            i += 1
            # skip raw table rows in text
            while i < len(lines) and (lines[i].startswith('Column 1') or '\t' in lines[i]):
                i += 1
            continue

        if line.startswith('TABLE II:'):
            if in_algo and algo_lines:
                add_code('\n'.join(algo_lines))
                algo_lines = []
                in_algo = False
            create_ieee_table(doc, headers_t2, rows_t2, [1.05, 0.75, 1.60], line)
            i += 1
            while i < len(lines) and (lines[i].startswith('Column 1') or '\t' in lines[i]):
                i += 1
            continue

        if line.startswith('TABLE III:'):
            if in_algo and algo_lines:
                add_code('\n'.join(algo_lines))
                algo_lines = []
                in_algo = False
            create_ieee_table(doc, headers_t3, rows_t3, [1.20, 0.55, 0.55, 0.55, 0.55], line)
            i += 1
            while i < len(lines) and (lines[i].startswith('Column 1') or '\t' in lines[i]):
                i += 1
            continue

        if line.startswith('TABLE IV:'):
            if in_algo and algo_lines:
                add_code('\n'.join(algo_lines))
                algo_lines = []
                in_algo = False
            create_ieee_table(doc, headers_t4, rows_t4, [1.00, 0.90, 0.75, 0.75], line)
            i += 1
            while i < len(lines) and (lines[i].startswith('Column 1') or '\t' in lines[i]):
                i += 1
            continue

        # Check Algorithm Title (when appearing as code block header)
        if line.startswith('Algorithm ') and ('Input:' in lines[i+1] if i+1 < len(lines) else False):
            if in_algo and algo_lines:
                add_code('\n'.join(algo_lines))
                algo_lines = []
            in_algo = True
            algo_lines.append(line)
            i += 1
            continue

        if in_algo:
            # Check if still inside algorithm
            if line.startswith('Input:') or line.startswith('width(') or line.startswith('Output:') or re.match(r'^\d+:\s*.*$', line):
                algo_lines.append(line)
                i += 1
                continue
            else:
                add_code('\n'.join(algo_lines))
                algo_lines = []
                in_algo = False
                # do not increment i, let loop process this line

        # Check H2
        if re.match(r'^[A-Z]\.\s+.*$', line):
            add_h2(line)
            i += 1
            continue

        # Check H3
        if re.match(r'^\d+\)\s+.*$', line):
            # check if it's a short title or an inline titled paragraph
            if ':' in line and len(line) > 50:
                # it's a paragraph with heading prefix
                add_smart_p(doc, line, indent=True)
            else:
                add_h3(line)
            i += 1
            continue

        # Check References
        if re.match(r'^\[\d+\]\s+.*$', line):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(2.0)
            p.paragraph_format.line_spacing = 1.0
            p.paragraph_format.left_indent = Inches(0.25)
            p.paragraph_format.first_line_indent = Inches(-0.25)

            r = p.add_run(line)
            r.font.name = 'Times New Roman'
            r.font.size = Pt(8.0)
            mark_no_proof(r)
            i += 1
            continue

        # Check Equation
        if re.search(r'\(\d+\)$', line) and not line.startswith('Where') and not line.startswith('Setting') and len(line) < 120:
            add_eq(line)
            i += 1
            continue

        # Check Bullet
        if line.startswith('•') or line.startswith('-'):
            add_bullet_p(doc, line)
            i += 1
            continue

        # Normal Body paragraph
        add_smart_p(doc, line, indent=True)
        i += 1

    if in_algo and algo_lines:
        add_code('\n'.join(algo_lines))

    doc.save(output_path)
    print(f"Saved document successfully to: {output_path}")

if __name__ == '__main__':
    text_file = 'user_clean.txt'
    out_file = r'C:\Users\r1tpr\.gemini\antigravity\scratch\compiler_paper\test_built_user_doc.docx'
    build_document(text_file, out_file, h1_gap=12.0)
