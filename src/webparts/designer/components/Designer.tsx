import { Stack } from "@fluentui/react";
import * as React from "react";
import styles from "./Designer.module.scss";
import { IDesignerProps } from "./IDesignerProps";

const Designer = (props: IDesignerProps): JSX.Element => {
    console.log(props);

    React.useEffect(() => {
        if (props.imageSrc) {
            const imageContainer = document.getElementById(
                "designer-image"
            ) as HTMLDivElement;
            imageContainer.style.visibility = "visible";
        }
    }, [props.imageSrc]);

    return (
        <Stack>
            <>
                <div className={styles.miniApp} id='mini-container' />
                <div className={styles.fullApp} id='full-container' />
            </>

            <img
                className={styles.designerImage}
                alt='Image generated by Designer'
                id='designer-image'
                src={props.imageSrc}
            />
        </Stack>
    );
};

export default Designer;
