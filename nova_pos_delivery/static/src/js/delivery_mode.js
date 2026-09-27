import { patch } from "@web/core/utils/patch";
import { ProductScreen } from "@point_of_sale/app/screens/product_screen/product_screen";

patch(ProductScreen.prototype, {
    setDeliveryMode(mode) {
        if (!this.currentOrder) {
            return;
        }
        this.currentOrder.update({ delivery_mode: mode });
    },

    get currentDeliveryMode() {
        return this.currentOrder?.delivery_mode || "pickup";
    },
});
