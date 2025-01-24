<template>
    <div>
        <v-btn class="mb-4 mr-2" density="compact" color="primary" @click="onClickHandler()">{{ ButtonText }}</v-btn>
        <br/>
    </div>
</template>

<script>
export default {
    props: {
        ButtonText: {
            type: String,
            required: true,
        },

        // density: {
        //     type: String,
        //     default: 'compact' | 'default' | 'comfortable'
        // },
        // color: {
        //     type: String,
        //     default: 'primary'
        // },
        clickHandler: {
            type: Function,
            required: true,
        }
    },

    methods: {
        onClickHandler() {
            this.clickHandler()
        }
    }
}
</script>

<!-- <style scoped>
.v-btn {
    min-width: 64px;
}
</style> -->