// Tablo dosya G/Ç Worker'ı: açma, yazma ve otomatik kayıt doğrulaması ana iş parçacığının dışında.
import { exposeIo } from "@officehere/core";
import { handlers } from "./ioHandlers";

exposeIo(handlers);
