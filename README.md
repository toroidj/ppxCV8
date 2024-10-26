Paper Plane xUI  V8 Script Module

.NET 用の Clear Script の V8 を利用して PPx 上で Javascript を実行する
PPx Module です。Clear Script は、WSH の JScript とある程度の
互換性を備えているので、いくつか修正することで WSH Script Module と
スクリプトを共有することができます。
また、.NET api が使用できます。

WSH 版との性能の比較としては、実行速度が jscript9.dll と同程度に
高速であり、最初の初期化に掛かる時間が秒単位になるほど遅い、
となります。
