# ETradeCSVToQFX

This script converts E\*Trade transactions in a CSV file into a QFX file for consumption by Quicken.
This script is useful to workaround the odd limitations in E\*Trade statement download functionality.
Eventhough E\*Trade allows transactions to be downloaded in both CSV and QFX formats,
there was a time that transactions older than 3 months could only be downloaded with the CSV format.
This odd QFX format limitation should not be confused with the limitation that neither format can contain more than a 3 month duration of transactions.

In order to use this script,
the user will have to download at least one transaction in a QFX file as well as the desired transactions in a CSV file.
The QFX file is used as a template for other non-transaction data, and its transactions are discarded during processing.

