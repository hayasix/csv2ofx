#!/usr/bin/env python3.5
# vim: set fileencoding=utf-8 fileformat=unix :

"""CSV to OFX converter.

Usage: {0} [options] PATH...

Options:
  -h, --help            show this help message and exit.
  -v, --version         show version.
  -f, --conf <conf>     read settings from CONF.
  --encoding <encoding>  specify encoding of CONF.
  -i, --issuer <issuer>  issuer defined as section in CONF.
  -z, --timezone <timezone>  timezone eg. GMT+0, JST-9, PST+8.
  --upper               coerce description (name of counterpart) to uppercase.

"""

import sys
import os
import csv
import datetime
import glob
import re
import unicodedata
from textwrap import dedent


__author__ = "HAYASI Hideki"
__email__ = "linxs@linxs.org"
__copyright__ = "Copyright (C) 2012 HAYASI Hideki <linxs@linxs.org>"
__license__ = "ZPL 2.1"
__version__ = "1.0.0a7"
__status__ = "Development"


REFMARK = unicodedata.lookup("REFERENCE MARK")

DEFAULT_CSV_ENCODING = "cp932"
DEFAULT_TIMEZONE = "JST-9"

HEADER = """\
OFXHEADER:100
DATA:OFXSGML
VERSION:102
SECURITY:NONE
ENCODING:UTF-8
CHARSET:CSUNICODE
COMPRESSION:NONE
OLDFILEUID:NONE
NEWFILEUID:NONE

<OFX>
 <SIGNONMSGSRSV1>
  <SONRS>
   <STATUS>
    <CODE>0
    <SEVERITY>INFO
   </STATUS>
   <DTSERVER>{datetime}
   <LANGUAGE>JPN
   <FI>
    <ORG>{cardname}
   </FI>
  </SONRS>
 </SIGNONMSGSRSV1>
 <CREDITCARDMSGSRSV1>
  <CCSTMTTRNRS>
   <TRNUID>0
   <STATUS>
    <CODE>0
    <SEVERITY>INFO
   </STATUS>
   <CCSTMTRS>
    <CURDEF>JPY
    <CCACCTFROM>
     <ACCTID>{cardnumber}
    </CCACCTFROM>
    <BANKTRANLIST>
     <DTSTART>{firstdate}
     <DTEND>{lastdate}
""".replace("\r\n", "\n")

TRANSACTION = """\
     <STMTTRN>
      <TRNTYPE>{transactiontype}
      <DTPOSTED>{datetime}
      <TRNAMT>{amount}
      <FITID>{fitid}
      <NAME>{description}
      <MEMO>{memo}
     </STMTTRN>
""".replace("\r\n", "\n")

FOOTER = """\
    </BANKTRANLIST>
    <LEDGERBAL>
     <BALAMT>{totalamount}
    </LEDGERBAL>
   </CCSTMTRS>
  </CCSTMTTRNRS>
 </CREDITCARDMSGSRSV1>
</OFX>
""".replace("\r\n", "\n")


def parse_fielddef(cols):
    """Build the reverse lookup table for field positions.

    cols        (str) comma-separated field names

    Returns a dict, each key of which is a field name and its associated
    value is the field position.
    """
    dic = dict()
    for i, col in enumerate(cols.split(",")):
        col = col.strip()
        if col in dic:
            if not isinstance(dic[col], list):
                dic[col] = [dic[col]]
            dic[col].append(i)
        elif col:
            dic[col] = i
        # else:  # if not col: continue
    return dic


def parse_date(s, tzinfo=None):
    """Parse a date string.

    s           (str) date string; format:
                    'YYYY/mm/dd' | 'YYYY-mm-dd' | 'YYYYmmdd'
    tzinfo      (Timezone)

    Returns a datetime.datetime with tzinfo as the timezone information.
    If tzinfo=None, a naive (timezone-less) datetime.dateme is returned.
    """
    if not isinstance(s, str):
        return None
    dt = None
    for fmt in ("%Y/%m/%d", "%Y-%m-%d", "%Y%m%d"):
        try:
            dt = datetime.datetime.strptime(s, fmt)
            break
        except ValueError:
            pass
    else:
        raise ValueError("illegal date format:" + s)
    if tzinfo:
        dt = dt.replace(tzinfo=tzinfo)
    return dt


class Timezone(datetime.tzinfo):

    """A concrete class of datetime.tzinfo.

    >>> Timezone('JST', -9).utcoffset()
    datetime.timedelta(0, 32400)

    """

    def __init__(self, tzname, utcoffset, dst=0):
        self._tzname = tzname
        # utcoffset should be in POSIX style; negative for eastern world.
        utcoffset = - utcoffset
        self._utcoffset = datetime.timedelta(hours=utcoffset)
        self._dst = datetime.timedelta(dst)

    def tzname(self, dt=None):
        return self._tzname

    def utcoffset(self, dt=None):
        return self._utcoffset

    def dst(self, dt=None):
        return self._dst


class Transaction(object):

    """A transaction record."""

    def __init__(self,
            date=datetime.date.today(),
            description="unknown",
            amount=0,  # > 0 to increase asset, < 0 to increase debt/capital
            category="unknown",
            tags=None,  # list of tags
            memo="",
            account="unknown",
            status="-",  # "-", "C" (cleared) or "R" (reconciled)
            tzinfo=None,
            ):
        self.date = date
        self.description = description
        self.amount = amount
        self.category = category
        self.tags = tags or []
        self.memo = memo
        self.account = account
        self.status = status
        self.tzinfo = tzinfo

    def __repr__(self):
        return "Transaction({dt}:{dsc}:{amt}:{cat}:{tag}:{mem}:{act}:{sts}:{tz})".format(
                dt=self.date.strftime("%Y-%m-%d"), dsc=self.description,
                amt=self.amount, cat=self.category, tag=",".join(self.tags),
                mem=self.memo, act=self.account, sts=self.status,
                tz=self.tzinfo or "")

    def __str__(self):
        return dedent("""\
                Date: {dt}
                Description: {dsc}
                Amount: {amt}
                Category: {cat}
                Tags: {tag}
                Memo: {mem}
                Account: {act}
                Status: {sts}
                Timezone: {tz}
                """).format(
                        dt=self.date.strftime("%Y-%m-%d"),
                        dsc=self.description,
                        amt=self.amount,
                        cat=self.category,
                        tag=",".join(self.tags),
                        mem=self.memo,
                        act=self.account,
                        sts=self.status,
                        tz=self.tzinfo or "")


class Journal(set):

    """A journal i.e. collection of transactions."""

    @staticmethod
    def ofxdatetime(dt):
        """Build a datetime string that complies with OFX standard.

        dt      (datetime.datetime)

        Returns a str.
        """
        if not dt.tzinfo:  # naive localtime
            return dt.strftime("%Y%m%d%H%M%S")
        return dt.strftime("%Y%m%d%H%M%S[{gmtoffset:+.2f}:{tzname}]").format(
                gmtoffset=dt.tzinfo.utcoffset().seconds / 3600.0,
                tzname=dt.tzname())

    def read_csv(self,
            pathname,
            accounttype="credit",
            cardnumber=None,
            cardname=None,
            header=None,
            fields=None,  # date, amount, description, memo, commission
            encoding=None,
            tzinfo=None,
            **option):
        """Read transactions from CSV file.

        pathname        (str) pathname of the source CSV file
        accounttype     (str) 'bank' | 'credit'
        cardnumber      (str) card number (16 digits or so)
                        (int) field position of card number (if header==True)
        cardname        (str) card name (card holder's name)
                        (int) field position of card name (if header==True)
        header          (bool) Read card number/name from the header
        fields          (str) comma-separated field names; a sequence of
                        'date', 'amount', 'description', 'memo' and
                        'commission'
        encoding        (str) encoding of the source CSV file
        tzinfo          (datetime.tzinfo) timezone for transactions
        **option        (dict) (ignored currently)

        Returns None.  To get transactions read, iterate over self.
        """
        fields = parse_fielddef(fields)
        # Read CSV header.
        encoding = encoding or DEFAULT_CSV_ENCODING
        with open(pathname, "r", encoding=encoding) as f:
            reader = csv.reader(f)
            if header:
                header = next(reader)
                cardnumber = header[cardnumber]
                cardname = header[cardname]
            elif header is not None:
                next(reader)  # Skip 1 line.
            # Read transactions.
            prev_date = datetime.datetime(2000, 1, 1)
            c = lambda f: line[fields[f]]
            n = lambda f: int(c(f).replace(",", "") or "0")
            for i, line in enumerate(reader):
                t = Transaction()
                try:
                    t.date = parse_date(c("date"))
                except ValueError:
                    t.date = prev_date
                t.date.replace(tzinfo=tzinfo)
                t.description = c("description")
                if accounttype == "bank":
                    t.amount = n("+amount") - n("-amount")
                else:  # accounttype == "credit":
                    t.amount = - n("amount")
                if "memo" in fields:
                    if isinstance(fields["memo"], (list, tuple)):
                        t.memo = ",".join(line[col] for col in fields["memo"])
                    else:
                        t.memo = c("memo")
                else:
                    t.memo = ""
                if "commission" in fields:
                    if not t.description or t.description.startswith(REFMARK):
                        continue
                    if not t.amount:
                        t.amount = - n("commission")
                t.fitid = i  # to overcome buggy OFX's
                self.add(t)
                prev_date = t.date
            self.cardnumber = cardnumber
            self.cardname = cardname
            self.datetime = datetime.datetime.now(tzinfo)

    def write_ofx(self, pathname, upper=False):
        """Write transactions as a OFX stream.

        pathname        (str) location to write transactions out
        upper           (bool) coerce description to uppercase

        Returns None.
        """
        normalize = lambda s: unicodedata.normalize("NFKC", s)
        xcase = lambda s: s.upper() if upper else s
        # Build OFX data.
        result = [HEADER.format(
                datetime=self.ofxdatetime(self.datetime),
                cardname=self.cardname,
                cardnumber=self.cardnumber,
                firstdate=self.ofxdatetime(min(t.date for t in self)),
                lastdate=self.ofxdatetime(max(t.date for t in self)),
                )]
        result.extend(TRANSACTION.format(
                transactiontype="CREDIT",
                datetime=self.ofxdatetime(t.date),
                amount=t.amount,
                fitid=t.fitid,
                description=xcase(normalize(t.description)),
                memo=normalize(t.memo),
                ) for t in sorted(self, key=lambda t: t.fitid))
        result.append(FOOTER.format(
                totalamount=sum(t.amount for t in self)))
        with open(pathname, "w", encoding="utf-8") as f:
            f.writelines(result)


def getencoding(path):
    """Detect encoding string from the leading two lines.

    path        (str) pathname of the source file

    Returns an encoding str or None.
    """
    coding = re.compile(r"coding[:=]\s*(\w)+")
    with open(path, encoding="ascii") as in_:
        for _ in (0, 1):
            try:
                mo = coding.search(in_.readline())
            except UnicodeDecodeError:
                continue
            if mo:
                return mo.group(0)
    return None


def gettimezone(timezone):
    """Get a Timezone.

    timezone        (str) timezone string

    Returns a Timezone.

    >>> gettimezone('JST-9').utcoffset()
    datetime.timedelta(0, 32400)

    """
    p = timezone.find("+")
    if p < 0:
        p = timezone.find("-")
    if p < 0:
        raise ValueError("illegal timezone format:" + timezone)
    return Timezone(timezone[:p].upper(), int(timezone[p:]))


def preprocess_btmucc(pathname):
    """Special preprocessor for the odd CSV files presented by BTMU.

    pathname        (str) pathname of the source CSV file

    Returns None.

    This function eliminate the extraordinariness in its header part.
    The original file is preserved but renamed with the additional suffix
    '.orig'.
    """
    if not pathname.lower().endswith(".csv"):
        return
    origname = pathname + ".orig"
    os.rename(pathname, origname)
    cr, lf, crlf = "\x0D", "\x0A", "\x0D\x0A"
    with open(origname, "r", encoding="cp932", newline=crlf) as in_, \
         open(pathname, "w", encoding="cp932") as out:
        sublines = in_.readline().rstrip().split(cr)
        out.write(sublines[-1] + crlf)
        out.write(in_.read())


def main(docstring):
    import docopt
    import configparser
    args = docopt.docopt(docstring.format(__file__), version=__version__)
    for k, v in args.items():
        setattr(args, k.lstrip("-"), v)
    args.conf = os.path.expanduser(args.conf)
    args.encoding = args.encoding or getencoding(args.conf) or "utf-8"
    conf = configparser.SafeConfigParser(dict(
            encoding=DEFAULT_CSV_ENCODING,
            timezone=DEFAULT_TIMEZONE,
            cardnumber=None,
            cardname=None,
            ))
    if args.encoding.lower().replace("_", "-") == "utf-8":
        args.encoding = "utf-8-sig"
    conf.read(args.conf, encoding=args.encoding)
    tz = args.timezone or conf.get("DEFAULT", "timezone")
    tzinfo = tz and gettimezone(tz) or None
    cardnumber = conf.get(args.issuer, "cardnumber")
    cardname = conf.get(args.issuer, "cardname")
    encoding = conf.get(args.issuer, "encoding")
    accounttype = (conf.get(args.issuer, "type") or "credit").lower()
    try:
        header = parse_fielddef(conf.get(args.issuer, "head"))
        # Read card number/name from CSV.
        # NB. Explicit cardnumber/cardname assignments take priority over
        # definitions in header line.
        if "cardnumber" in header:
            cardnumber = cardnumber or header["cardnumber"]
        if "cardname" in header:
            cardname = cardname or header["cardname"]
    except configparser.NoOptionError:
        header = None
    body = conf.get(args.issuer, "body")

    for path in args.PATH:
        if "*" in path or "?" in path:
            filelist = glob.glob(path)
        else:
            filelist = [path]
        for in_ in filelist:
            if not in_.lower().endswith(".csv"):
                raise ValueError("only CSV files are acceptable")
            if args.issuer.lower() == "btmucc":
                preprocess_btmucc(in_)
            out = in_[:-4] + ".ofx"
            journal = Journal()
            journal.read_csv(in_,
                    accounttype=accounttype,
                    cardnumber=cardnumber, cardname=cardname,
                    header=header, fields=body, encoding=encoding,
                    tzinfo=tzinfo)
            journal.write_ofx(out, upper=args.upper)


if __name__ == "__main__":
    main(__doc__)
