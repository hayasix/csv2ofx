#!/usr/bin/env python3
# vim: set fileencoding=utf-8 fileformat=unix :

"""CSV to OFX converter.

Usage: {script} [options] [PATH...]

Options:
  -h, --help                show this help message and exit.
  -v, --version             show version.
  -f, --conf <conf>         read settings from CONF.
  -i, --issuer <issuer>     issuer defined as section in CONF.
  -a, --amazon <file>       specify Amazon.co.jp order history file
  -s, --subst <file>        specify user-defined memo substitution table
  -z, --timezone <tz>       timezone eg. GMT+0, JST-9, PST+8.
  -l, --show-issuers        show issuer list
  --start-date YYYYMMDD     start date to report
  --end-date YYYYMMDD       end date to report
  --encoding <encoding>     specify encoding of CONF.
  --upper                   coerce description to uppercase.
  --debug                   equivalent to --loglevel DEBUG
  --loglevel LEVEL          set loglevel to LEVEL [default: WARNING]
  --logfile FILE            write log into FILE [default: {logfile}]

"""

import sys
import os
from pathlib import Path
import csv
import datetime
import glob
import re
import unicodedata
from textwrap import dedent
from configparser import ConfigParser, NoOptionError
import logging


__author__ = "HAYASHI Hideki"
__email__ = "hideki@hayasix.com"
__copyright__ = "Copyright (C) 2012 HAYASHI Hideki <hideki@hayasix.com>"
__license__ = "ZPL 2.1"
__version__ = "1.0.0b8"
__status__ = "Development"


REFMARK = unicodedata.lookup("REFERENCE MARK")
UTF8BOM = b"\xef\xbb\xbf"  # "\ufeff"

CONFIGS = ["~/.config/csv2ofx/config", "~/csv2ofx.ini"]
PARAMETERS = ["encoding", "timezone", "type", "cardnumber", "cardname",
              "skip", "head", "body"]
DEFAULTS = dict(encoding="cp932", timezone="JST-9")

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


if sys.version_info < (3, 8):
    raise RuntimeError("Python < 3.8 is not supported")


def normalize(s):
    s = re.sub(r"([ｧ-ﾜ])-", "\\1\uff70", s)
    s = re.sub(r"([ァ-ワぁ-わ])−", "\\1\\u30fc", s)
    return unicodedata.normalize("NFKC", s)


def parse_fielddef(cols):
    """Build the reverse lookup table for field positions.

    Parameters
    ----------
    cols : str
        comma-separated field names

    Returns
    -------
    dict
        each key of which is a field name and its associated value is
        the field position.
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

    s : str
        date string; format: 'YYYY/mm/dd' | 'YYYY-mm-dd' | 'YYYYmmdd'
    tzinfo : Timezone

    Returns
    -------
    datetime.datetime
        date and time with tzinfo as the timezone information

    If tzinfo=None, a naive (timezone-less) datetime.dateme is returned.
    """
    if isinstance(s, (datetime.date, datetime.datetime)):
        return s
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
        return "Transaction(" + ":".join((
                self.date.strftime("%Y-%m-%d"),
                self.description,
                str(self.amount),
                self.category,
                ",".join(self.tags),
                self.memo,
                self.account,
                self.status,
                self.tzinfo or "",
                )) + ")"

    def __str__(self):
        return dedent(f"""\
                Date: {self.date.strftime("%Y-%m-%d")}
                Description: {self.description}
                Amount: {self.amount}
                Category: {self.category}
                Tags: {",".join(self.tags)}
                Memo: {self.memo}
                Account: {self.account}
                Status: {self.status}
                Timezone: {self.tzinfo or ""}
                """)


def detect_encoding(path):
    """Detect the encoding of a text file."""
    pat = re.compile(b"^#.*coding[:=]\s*([\w\-]+)", re.I)
    with open(path, "rb") as in_:
        if in_.read(3) == UTF8BOM: return "utf-8-sig"
        in_.seek(0)
        mo = [pat.match(in_.readline()), pat.match(in_.readline())]
    return ([m.group(1) for m in mo if m] or [b"utf-8"])[0].decode()


class Journal(set):

    """A journal i.e. collection of transactions."""

    @staticmethod
    def ofxdatetime(dt):
        """Build a datetime string that complies with OFX standard.

        Parameters
        ----------
        dt : datetime.datetime

        Returns
        -------
        str
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
            skip=None,
            fields=None,  # date, amount, description, memo, commission
            encoding=None,
            tzinfo=None,
            amazon=None,
            subst=None,
            **option):
        """Read transactions from CSV file.

        Parameters
        ----------
        pathname : str
            pathname of the source CSV file
        accounttype : str
            'bank' | 'credit'
        cardnumber : str | int
            (str) card number (16 digits or so)
            (int) field position of card number (if header==True)
        cardname : str | int
            (str) card name (card holder's name)
            (int) field position of card name (if header==True)
        header : int | bool
            (int) header lines to skip
            (bool) read card number/name from the header
        skip : int | None
        fields : str
            comma-separated field names;
            a sequence of 'date', 'amount', 'description', 'memo', 'commission';
            '+amount' and '-amount', and 'date?' are also available
        encoding : str
            encoding of the source CSV file
        tzinfo : datetime.tzinfo
            timezone for transactions
        amazon : str
            Amazon.co.jp order history
        subst : str
            User-defined memo substitution table
        **option : dict
            (ignored currently)

        Returns
        -------
        None

        To get transactions read, iterate over self.
        """
        if amazon:
            az = AmazonJournal()
            az.read_csv(amazon)
        if subst:
            enc = detect_encoding(subst)
            # Setup the memo substitution table.
            substdic = dict()
            with open(subst, "r", encoding=enc) as in_:
                for line in in_:
                    if line.startswith("#"): continue
                    if "=" not in line: continue
                    k, v = line.strip().split("=", 1)
                    substdic[k] = v
        fields = parse_fielddef(fields)
        datefield = [f for f in fields if f in ("date", "date?")][0]
        with open(pathname, "r", encoding=encoding) as f:
            reader = csv.reader(f)
            # Read CSV header.
            if skip:
                for _ in range(skip):
                    next(reader)  # Skip N lines.
            if header:
                header = next(reader)
                if isinstance(cardnumber, (int, float)):
                    cardnumber = header[cardnumber]
                if isinstance(cardname, (int, float)):
                    cardname = header[cardname]
            elif header is not None:  # ''
                next(reader)  # Skip 1 line.
            # Read transactions.
            prev_date = datetime.datetime(2000, 1, 1)
            def c(f, defval=None):
                if f.endswith("?"):
                    return normalize(line[fields[f.rstrip("?")]]) or defval
                return normalize(line[fields[f]])
            def n(f):
                return int(normalize(c(f)).replace(",", "") or "0")
            for i, line in enumerate(reader):
                logging.debug(f">>> {line}")
                t = Transaction()
                try:
                    t.date = parse_date(c(datefield, defval=prev_date))
                except ValueError:
                    continue
                t.date.replace(tzinfo=tzinfo)
                t.description = c("description")
                try:
                    t.amount = abs(n("+amount")) - abs(n("-amount"))
                except KeyError:
                    try:
                        t.amount = n("amount")
                        if accounttype == "credit":
                            t.amount *= -1
                    except KeyError:
                        t.amount = n("-amount")
                        assert accounttype == "credit"
                if "memo" in fields:
                    if isinstance(fields["memo"], (list, tuple)):
                        t.memo = ",".join(line[col] for col in fields["memo"]
                                                            if line[col])
                    else:
                        t.memo = c("memo")
                else:
                    t.memo = ""
                t.description = re.sub(" +", " ", t.description)
                t.memo = re.sub(" +", " ", t.memo)
                if amazon and t.description == "AMAZON.CO.JP":
                    txns = az.search(date=t.date.date(), amount=-t.amount)
                    if len(txns) != 1:
                        sys.stderr.write(
                            "W: multiple or no card charges found "
                            "for date={}, amount={}\n".format(
                                t.date.date(), -t.amount))
                    else:
                        t.memo = txns[0][1]
                # Fix memo using the user-defined substitution table.
                if subst:
                    for k, v in substdic.items():
                        t.memo.replace(k, v)
                # Remove duplicate description from memo.
                dlen = len(t.description)
                if (t.memo[:dlen] == t.description and
                        t.memo[dlen:].startswith(",")):
                    t.memo = t.memo[dlen + 1:]
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

    def write_ofx(self, pathname, upper=False,
                    start_date=None, end_date=None):
        """Write transactions as a OFX stream.

        Parameters
        ----------
        pathname : str
            location to write transactions out
        upper : bool
            coerce description to uppercase
        start_date : datetime.date | '' | None
            write out transactions on or after this date
        end_date : datetime.date | '' | None
            write out transactions on or before this date

        Returns
        -------
        None

        Notes
        -----
        No file will be created if no transactions are recorded.
        """
        if len(self) < 1: return
        if start_date:
            start_date = parse_date(start_date)
        if end_date:
            end_date = parse_date(end_date)
        def in_period(d):
            return ((not start_date or start_date <= d)
                    and (not end_date or d <= end_date))
        xcase = lambda s: s.upper() if upper else s
        # Build OFX data.
        subset = [t for t in self if t.amount and in_period(t.date)]
        subset.sort(key=lambda t: t.fitid)
        result = [HEADER.format(
                datetime=self.ofxdatetime(self.datetime),
                cardname=self.cardname,
                cardnumber=self.cardnumber,
                firstdate=self.ofxdatetime(min(t.date for t in subset)),
                lastdate=self.ofxdatetime(max(t.date for t in subset)),
                )]
        result.extend(TRANSACTION.format(
                transactiontype="CREDIT" if 0 <= t.amount else "DEBIT",
                datetime=self.ofxdatetime(t.date),
                amount=abs(t.amount),
                fitid=t.fitid,
                description=xcase(normalize(t.description)),
                memo=normalize(t.memo),
                ) for t in subset)
        result.append(FOOTER.format(
                totalamount=sum(t.amount for t in subset)))
        with open(pathname, "w", encoding="utf-8") as f:
            f.writelines(result)


class AmazonOrderItem(dict):

    """Single item with name, unit price, quantity, etc."""

    def __init__(self, row):
        name = row["商品名"]
        description = row["付帯情報"].replace("　", " ").split(" ", 1)[0]
        self.update(dict(
            orderid = row["注文番号"],
            name = row["商品名"],
            description = row["付帯情報"],
            price = row["価格"],
            quantity = row["個数"],
            # Following data will NOT be fed for あわせ買い対象商品
            amount = row["商品小計"],  # = price * quantity
            ))

    def __str__(self):
        return self["name"]


class AmazonOrderBatch(list):

    """List of item-sets.

    An item-set is a list of an item and its add-on items.
    So, AmazonOrderBatch is a list of lists.
    """

    def __init__(self, row):
        self.append([AmazonOrderItem(row)])

    def add(self, row):
        self[-1].append(AmazonOrderItem(row))

    @staticmethod
    def omit(s, width=40):
        def charwidth(c):
            return (1, 2)[unicodedata.east_asian_width(c) in "FWA"]
        cw = list(map(charwidth, normalize(s).replace("　", " ")))
        if sum(cw) <= width: return s
        t = -1
        maxlen = int((width - 2) / 2)
        while sum(cw[t - 1:]) <= maxlen: t -= 1
        h = 1
        maxlen = width - sum(cw[t:]) - 2
        while sum(cw[:h + 1]) <= maxlen: h += 1
        return s[:h] + "." * (width - sum(cw[:h]) - sum(cw[t:])) + s[t:]


    def omitted(self):
        return ";".join(",".join(self.omit(item["name"]) for item in itemset)
                        for itemset in self)

    def __str__(self):
        return ";".join(",".join(item["name"] for item in itemset)
                        for itemset in self)


class AmazonOrder(list):

    """List of AmazonOrderBatches. """

    def __init__(self, orderid):
        self.orderid = orderid
        self.charges = []

    def __str__(self):
        return ";;".join(map(str, self))

    @staticmethod
    def ccc(row):
        return (row["クレカ請求日"], row["クレカ請求額"], row["クレカ種類"])

    def add_charge(self, row):
        self.charges.append(self.ccc(row))

    def add_row(self, row):
        name = row["商品名"]
        if name in ("（注文全体）", "（割引）", "（配送料・手数料）",
                    "（Amazonポイント）"):
            return
        if name == "（クレジットカードへの請求）":
            ccc = self.ccc(row)
            if ccc in self.charges:  # Charge records sometimes duplicate!
                self.charges.remove(ccc)
            self.add_charge(row)
            return
        if row["クレカ請求額"]:  # Only for digitally sold items.
            row["クレカ請求日"] = row["注文日"]
            self.append(AmazonOrderBatch(row))
            self.add_charge(row)
            return
        if row["商品小計"] != "":
            self.append(AmazonOrderBatch(row))
        else:  # Add-on items.
            self[-1].add(row)

    def order_charge_pairs(self):
        if len(self.charges) == 1:
            return [(self, self.charges[0])]
        return [(str(item), self.charges[i]) for i, item in enumerate(self)]


class AmazonJournal(dict):

    def read_csv(self, pathname):
        with open(pathname, "r", encoding="utf-8-sig") as in_:
            reader = csv.DictReader(in_)
            for row in reader:
                order_id = row["注文番号"]
                if order_id not in self:
                    self[order_id] = order = AmazonOrder(order_id)
                order.add_row(row)

    def search(self, date=None, amount=None):
        if date and not isinstance(date, (tuple, list)):
            date = (date - datetime.timedelta(days=1),
                    date + datetime.timedelta(days=2))
        result = []
        for order in self.values():
            for i, charge in enumerate(order.charges):
                chargedate = parse_date(charge[0]).date()
                chargeamount = int(charge[1])
                if ((date is None or date[0] <= chargedate <= date[1]) and
                        (amount is None or chargeamount == amount)):
                    result.append((order.orderid, order[i].omitted()))
        return result


def getencoding(path):
    """Detect encoding string from the leading two lines.

    Parameters
    ----------
    path : str
        pathname of the source file

    Returns
    -------
    str | None
        encoding specified in file
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

    Parameters
    ----------
    timezone : str
        timezone string

    Returns
    -------
    Timezone

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

    Parameters
    ----------
    pathname : str
        pathname of the source CSV file

    Returns
    -------
    None

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


def getparams(conf: ConfigParser, issuer: str, baselist=None) -> dict:
    params = dict()
    if conf.has_option(issuer, "include"):
        base = conf.get(issuer, "include").strip("[] ")
        if not baselist: baselist = list()
        if base == issuer or base in baselist:
            raise ValueError(f"include=['{base}'] causes an infinite loop")
        baselist.append(base)
        params.update(getparams(conf, base, baselist=baselist))
    for key in PARAMETERS:
        if not conf.has_option(issuer, key): continue
        params[key] = conf.get(issuer, key)
        if key != "head": continue
        # Read card number/name from CSV header.
        # NB. Explicit cardnumber/cardname assignments take priority over
        # definitions in header line.
        header = parse_fielddef(conf.get(issuer, "head"))
        for k in ("cardnumber", "cardname"):
            params.setdefault(k, header.get(k, ""))
    return params


def expandpath(path: str) -> str:
    return os.path.abspath(os.path.expanduser(os.path.expandvars(path)))


def findconf(path: str) -> str:
    for p in [path] + CONFIGS:
        if p and os.path.exists(p := expandpath(p)): return p
    raise ValueError("no configuration file")


def main(args: dict):
    for k, v in args.items():
        setattr(args, k.lstrip("-").replace("-", "_"), v)
    args.conf = findconf(args.conf)
    args.encoding = args.encoding or getencoding(args.conf) or "utf-8"
    conf = ConfigParser()
    if args.encoding.lower().replace("_", "-") == "utf-8":
        args.encoding = "utf-8-sig"
    conf.read(args.conf, encoding=args.encoding)
    if args.show_issuers:
        print("\n".join(sorted(s for s in conf if s != "DEFAULT")))
        return
    if not conf.has_section(args.issuer):
        raise ValueError(f"can't find issuer '{args.issuer}'")
    params = getparams(conf, args.issuer)
    for k, v in DEFAULTS.items(): params.setdefault(k, v)
    tz = args.timezone or params.get("timezone")
    tzinfo = gettimezone(tz) if tz else None
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
            def _(k): return params.get(k)
            journal.read_csv(in_,
                    accounttype=_("type"),
                    cardnumber=_("cardnumber"), cardname=_("cardname"),
                    header=_("header"), skip=int(_("skip") or "0"),
                    fields=_("body"), encoding=_("encoding"),
                    tzinfo=tzinfo, amazon=args.amazon, subst=args.subst)
            journal.write_ofx(out, upper=args.upper,
                    start_date=args["--start-date"],
                    end_date=args["--end-date"])


def setlogger(filename=None, loglevel=logging.WARNING, force=False):
    filename = Path(filename or __file__ or "current.log").with_suffix(".log")
    logging.basicConfig(filename=str(filename),
            format="%(asctime)s:%(levelname)s:%(module)s:%(message)s",
            encoding="utf-8-sig",
            level=loglevel,
            force=force)
    import atexit; atexit.register(logging.shutdown)


def __main__():
    """Support for logging."""
    from docopt import docopt
    args = docopt(__doc__.format(script=(_ := Path(__file__)).name,
                                 logfile=_.with_suffix(".log").name),
                  version=__version__)
    loglevel = logging.DEBUG if args["--debug"] else args["--loglevel"]
    setlogger(filename=args["--logfile"], loglevel=loglevel)
    import atexit; atexit.register(logging.shutdown)
    try:
        return main(args)
    except Exception:
        import traceback
        loglines = traceback.format_exc().splitlines()
        logging.error(loglines[-1])
        for logline in loglines[:-1]: logging.error(logline)
        raise


if __name__ == "__main__":
    sys.exit(__main__())
