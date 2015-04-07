import zipfile
import io
import re

from lxml import etree

import click

def copyfile(ininfo, inzip, outzip):
    print("Copying {} verbatim".format(ininfo.filename))
    outzip.writestr(ininfo, inzip.open(ininfo).read())

def modsheet(ininfo, cellchanges, inzip, outzip):
    print("Opening {} to make {} change{}".format(
        ininfo.filename,
        len(cellchanges),
        "s" if len(cellchanges) != 1 else "",
    ))
    xmldata = etree.parse(inzip.open(ininfo))

    # Because all the xml in these documents are namespaced, we have to
    # set up an explicit prefix to use that points to the default namespace.
    # The nsmap attribute one reason we have to use the third party
    # lxml library instead of the standard ElementTree library - there is no
    # way to get declared namespace info out of Python's ElementTree other
    # than manually parsing things with string manipulation.
    nsmap = xmldata.getroot().nsmap
    nsmap['d'] = nsmap[None]
    del nsmap[None]

    for cell, contents in cellchanges:
        # cell is an excel address e.g. "C12".
        # Search the xml tree for the appropriate element and change its value
        node = xmldata.find("/d:sheetData[1]/d:row/d:c[@r='{}']/d:v".format(
            cell
        ), nsmap)
        if node is None:
            raise RuntimeError("No such cell element with property r='{}' "
                               "found.".format(cell))
        print("  Modifying contents of cell {} at {}".format(
            node.getparent().get("r"),
            re.sub(r"({[^}]+})", "", xmldata.getelementpath(node))
        ))
        node.text = contents

    # Now go and recompute all formulas. Actually, this just deletes the
    # value element from any cell elements that have a formula element
    for elem in xmldata.findall("/d:sheetData[1]/d:row/d:c[d:f]", nsmap):
        value = elem.find("./d:v", nsmap)
        if value is not None:
            print("  Removing value for formula cell {} at {}".format(
                elem.get("r"),
                re.sub(r"({[^}]+})", "", xmldata.getelementpath(value)),
            ))
            elem.remove(value)

    out = io.BytesIO()
    xmldata.write(out)
    outzip.writestr(ininfo, out.getvalue())

def modchart(ininfo, inzip, outzip):
    print("Removing caches from chart {}".format(ininfo.filename))
    xmldata = etree.parse(inzip.open(ininfo))
    nsmap = xmldata.getroot().nsmap
    # The chart xml doesn't have a default namespace, but every element has
    # an explicit namespace

    # Delete any "numCache" elements from the tree
    for numcache_element in xmldata.findall("//c:numCache", nsmap):
        print("  Removing element at {}".format(xmldata.getpath(
            numcache_element)))
        numcache_element.getparent().remove(numcache_element)

    out = io.BytesIO()
    xmldata.write(out)
    outzip.writestr(ininfo, out.getvalue())

@click.command()
@click.argument("infile", type=click.Path(readable=True))
@click.argument("sheet")
@click.argument("cell")
@click.argument("content")
@click.argument("outfile", type=click.Path(writable=True))
def main(infile, sheet, cell, content, outfile):
    inzip = zipfile.ZipFile(infile, mode="r")
    outzip = zipfile.ZipFile(outfile, mode="w")
    for item in inzip.infolist():
        if item.filename == "xl/worksheets/{}.xml".format(sheet):
            modsheet(item, [(cell, content)], inzip, outzip)
        elif re.match(r"xl/worksheets/[^/]+\.xml", item.filename):
            modsheet(item, [], inzip, outzip)
        elif re.match(r"xl/charts/[^/]+\.xml", item.filename):
            modchart(item, inzip, outzip)
        else:
            copyfile(item, inzip, outzip)

    outzip.close()

if __name__ == "__main__":
    main()