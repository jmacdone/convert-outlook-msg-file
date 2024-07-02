import io
import email.message
import email.parser
from email.utils import formataddr, formatdate
import logging
import re

import compoundfiles
import html2text
from rtfparse.parser import Rtf_Parser
from rtfparse.renderers.html_decapsulator import HTML_Decapsulator

from .properties import parse_properties
from .attachments import process_attachment

logger = logging.getLogger(__name__)


def load(filename_or_stream):
    with compoundfiles.CompoundFileReader(filename_or_stream) as doc:
        doc.rtf_attachments = 0
        return load_message_stream(doc.root, True, doc)


def set_msg_headers_from_metadata(msg, props):
    # Construct common headers from metadata.

    try:
        msg["Date"] = formatdate(props["MESSAGE_DELIVERY_TIME"].timestamp())
    except KeyError:
        pass

    try:
        sender = props["SENDER_NAME"]
        try:
            representing = props["SENT_REPRESENTING_NAME"]
            if sender != representing:
                sender += f" ({representing})"
        except KeyError:
            pass
        msg["From"] = formataddr((props["SENDER_NAME"], ""))
    except KeyError:
        pass

    header_to_property = {
        "To": "DISPLAY_TO",
        "CC": "DISPLAY_CC",
        "BCC": "DISPLAY_BCC",
        "Subject": "Subject"
    }
    for header_name, property_name in header_to_property.items():
        try:
            value = props[property_name]
        except KeyError:
            continue
        if not value:
            continue
        msg[header_name] = value

def set_msg_headers_from_transport(msg, props):
    """Add the raw headers, if known."""
    # Get the string holding all of the headers.
    try:
       headers = props["TRANSPORT_MESSAGE_HEADERS"]
    except KeyError:
        return

    if isinstance(headers, bytes):
        headers = headers.decode("utf-8")

    # Remove content-type header because the body we can get this
    # way is just the plain-text portion of the email and whatever
    # Content-Type header was in the original is not valid for
    # reconstructing it this way.
    headers = re.sub(r"Content-Type: .*(\n\s.*)*\n", "", headers, flags=re.I)

    # Parse them.
    headers = email.parser.HeaderParser(policy=email.policy.default).parsestr(
        headers
    )

    # Copy them into the message object.
    for header, value in headers.items():
        try:
            msg[header] = value
        except ValueError:
            #trying to set an already-set single-value header
            logger.debug("Forcing %s: %s to %s", header, msg[header], value)
            del(msg[header])
            msg[header] = value


def load_message_stream(entry, is_top_level, doc):
    # Load stream data.
    props = parse_properties(entry["__properties_version1.0"], is_top_level, entry, doc)

    # Construct the MIME message....
    msg = email.message.EmailMessage()
    set_msg_headers_from_metadata(msg, props)
    set_msg_headers_from_transport(msg, props)

    # Add a plain text body from the BODY field.
    has_body = False
    if "BODY" in props:
        body = props["BODY"]
        if isinstance(body, str):
            msg.set_content(body, cte="quoted-printable")
        else:
            msg.set_content(body, maintype="text", subtype="plain", cte="8bit")
        has_body = True

    # Add a HTML body from the RTF_COMPRESSED field.
    if "RTF_COMPRESSED" in props:
        # Decompress the value to Rich Text Format.
        import compressed_rtf

        rtf = props["RTF_COMPRESSED"]
        rtf = compressed_rtf.decompress(rtf)

        # Try rtfparse to de-encapsulate HTML stored in a rich
        # text container.
        try:
            rtf_blob = io.BytesIO(rtf)
            parsed = Rtf_Parser(rtf_file=rtf_blob).parse_file()
            html_stream = io.StringIO()
            HTML_Decapsulator().render(parsed, html_stream)
            html_body = html_stream.getvalue()

            if not has_body:
                # Try to convert that to plain/text if possible.
                text_body = html2text.html2text(html_body)
                msg.set_content(text_body, subtype="text", cte="quoted-printable")
                has_body = True

            if not has_body:
                msg.set_content(html_body, subtype="html", cte="quoted-printable")
                has_body = True
            else:
                msg.add_alternative(html_body, subtype="html", cte="quoted-printable")

        # If that fails, just attach the RTF file to the message.
        except Exception:
            doc.rtf_attachments += 1
            fn = "messagebody_{}.rtf".format(doc.rtf_attachments)

            if not has_body:
                msg.set_content(
                    "<no plain text message body --- see attachment {}>".format(fn),
                    cte="quoted-printable",
                )
                has_body = True

            # Add RTF file as an attachment.
            msg.add_attachment(rtf, maintype="text", subtype="rtf", filename=fn)

    if not has_body:
        msg.set_content("<no message body>", cte="quoted-printable")

    # # Copy over string values of remaining properties as headers
    # # so we don't lose any information.
    # for k, v in props.items():
    #   if k == 'RTF_COMPRESSED': continue # not interested, save output
    #   msg[k] = str(v)

    # Add attachments.
    for stream in entry:
        if stream.name.startswith("__attach_version1.0_#"):
            try:
                process_attachment(msg, stream, doc)
            except KeyError as e:
                logger.error("Error processing attachment {} not found".format(str(e)))
                continue

    return msg
